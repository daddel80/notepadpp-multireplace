// This file is part of MultiReplace.
//
// MultiReplace is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 2 of the License, or
// (at your option) any later version.

#include "EcmdParser.h"

namespace MultiReplaceEngine {

namespace {

struct Cursor {
    const std::string& src;
    std::size_t        pos = 0;

    bool atEnd() const noexcept { return pos >= src.size(); }
    char peek(std::size_t off = 0) const noexcept {
        return pos + off < src.size() ? src[pos + off] : '\0';
    }
};

inline bool isAsciiSpace(char c) noexcept {
    return c == ' ' || c == '\t' || c == '\n' || c == '\r' || c == '\v' || c == '\f';
}

inline bool isIdentStart(char c) noexcept {
    return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == '_';
}

inline bool isIdentCont(char c) noexcept {
    return isIdentStart(c) || (c >= '0' && c <= '9');
}

// Skip whitespace and ExprTk-style comments. Returns false on an
// unterminated block comment so callers can report the start position;
// true otherwise. Forward-declared because the comment/string helpers
// below are used by skipTrivia.
bool skipTrivia(Cursor& c, std::size_t& commentStartOut) noexcept;

// Read an identifier [_A-Za-z][_A-Za-z0-9]*. Returns empty when the
// cursor isn't on an identifier start.
std::string readIdentifier(Cursor& c) {
    if (!isIdentStart(c.peek())) return {};
    const std::size_t start = c.pos;
    ++c.pos;
    while (!c.atEnd() && isIdentCont(c.peek())) ++c.pos;
    return c.src.substr(start, c.pos - start);
}

// Token-aware byte advance over the body looking for the closing `end`
// keyword. Handles ExprTk string literals ('...' with '' as escape),
// /* */ and // comments, and skips identifiers atomically so `end`
// inside `extended` doesn't trigger a false match. Returns the offset
// of the `end` keyword on success, or SIZE_MAX on a structural failure
// (unterminated string / comment, or EOF before `end`).
constexpr std::size_t kNotFound = static_cast<std::size_t>(-1);

// Skip a line comment (we're at "//"). Advances past the newline-or-EOF.
void skipLineComment(Cursor& c) noexcept {
    c.pos += 2;
    while (!c.atEnd() && c.peek() != '\n') ++c.pos;
}

// Skip a block comment (we're at "/*"). Returns false on unterminated;
// errPosOut then holds the start offset for the diagnostic.
bool skipBlockComment(Cursor& c, std::size_t& errPosOut) noexcept {
    const std::size_t start = c.pos;
    c.pos += 2;
    while (!c.atEnd()) {
        if (c.peek() == '*' && c.peek(1) == '/') {
            c.pos += 2;
            return true;
        }
        ++c.pos;
    }
    errPosOut = start;
    return false;
}

// Skip an ExprTk single-quoted string (we're at the opening quote).
// Doubled '' inside is the accepted escape per ExprTk docs section 13.
// Returns false on unterminated; errPosOut then holds the start offset.
bool skipStringLiteral(Cursor& c, std::size_t& errPosOut) noexcept {
    const std::size_t start = c.pos;
    ++c.pos;
    while (!c.atEnd()) {
        if (c.peek() == '\'') {
            if (c.peek(1) == '\'') { c.pos += 2; continue; }
            ++c.pos;
            return true;
        }
        ++c.pos;
    }
    errPosOut = start;
    return false;
}

// Definition of the forward-declared skipTrivia, now that the comment
// helpers above are in scope.
bool skipTrivia(Cursor& c, std::size_t& commentStartOut) noexcept {
    for (;;) {
        while (!c.atEnd() && isAsciiSpace(c.peek())) ++c.pos;
        if (c.atEnd()) return true;
        if (c.peek() == '/' && c.peek(1) == '/') { skipLineComment(c); continue; }
        if (c.peek() == '/' && c.peek(1) == '*') {
            if (!skipBlockComment(c, commentStartOut)) return false;
            continue;
        }
        return true;
    }
}

std::size_t findFunctionEnd(const std::string& src, std::size_t from,
                            std::string& errOut, std::size_t& errPosOut)
{
    Cursor c{src, from};
    while (!c.atEnd()) {
        const char ch = c.peek();

        if (isAsciiSpace(ch)) { ++c.pos; continue; }

        if (ch == '/' && c.peek(1) == '/') { skipLineComment(c); continue; }

        if (ch == '/' && c.peek(1) == '*') {
            if (!skipBlockComment(c, errPosOut)) {
                errOut = "Unterminated block comment";
                return kNotFound;
            }
            continue;
        }

        if (ch == '\'') {
            if (!skipStringLiteral(c, errPosOut)) {
                errOut = "Unterminated string literal";
                return kNotFound;
            }
            continue;
        }

        if (isIdentStart(ch)) {
            const std::size_t identStart = c.pos;
            const std::string ident = readIdentifier(c);
            if (ident == "end") return identStart;
            // 'function' inside a body means the user forgot `end`
            // before starting the next definition - much friendlier
            // error than letting the parser hit EOF on the outer loop.
            if (ident == "function") {
                errOut = "Unexpected 'function' inside body - missing 'end'?";
                errPosOut = identStart;
                return kNotFound;
            }
            continue;
        }

        // Everything else (operators, braces, parens, digits) is one
        // byte we don't need to interpret - the only structural
        // delimiters that matter to us are strings, comments, and
        // identifiers, all handled above.
        ++c.pos;
    }

    errOut = "Function body has no matching 'end'";
    errPosOut = from;
    return kNotFound;
}

// Parse a single parameter: identifier with optional ": T" or ": S".
// Cursor must be positioned on the identifier start; on success it
// stops at the byte after the type letter (or after the identifier
// when no annotation is present). Trailing whitespace is NOT consumed.
bool parseOneParam(Cursor& c, EcmdParser::ParamDef& out,
                   std::string& err, std::size_t& errPos)
{
    if (!isIdentStart(c.peek())) {
        err = "Expected parameter name";
        errPos = c.pos;
        return false;
    }
    out.sourcePos = c.pos;
    out.name = readIdentifier(c);

    std::size_t commentErr = 0;
    if (!skipTrivia(c, commentErr)) {
        err = "Unterminated block comment";
        errPos = commentErr;
        return false;
    }

    if (c.peek() != ':') {
        out.type = EcmdParser::ValueType::Scalar;
        return true;
    }
    ++c.pos;
    if (!skipTrivia(c, commentErr)) {
        err = "Unterminated block comment";
        errPos = commentErr;
        return false;
    }

    const std::size_t typePos = c.pos;
    if (c.atEnd()) {
        err = "Expected type letter after ':'";
        errPos = typePos;
        return false;
    }
    const char t = c.peek();
    // A structural character right after ':' means the user forgot the
    // type letter entirely - point that out specifically.
    if (t == ',' || t == ')' || t == ':') {
        err = "Expected type letter after ':'";
        errPos = typePos;
        return false;
    }
    if (t == 'T')      out.type = EcmdParser::ValueType::Scalar;
    else if (t == 'S') out.type = EcmdParser::ValueType::String;
    else {
        // Specifically diagnose the common lowercase mistake.
        if (t == 's' || t == 't') {
            err = "Type annotation must be uppercase (S or T), got '";
            err += t;
            err += '\'';
        } else {
            err = "Unknown type annotation '";
            err += t;
            err += "' - expected S or T";
        }
        errPos = typePos;
        return false;
    }
    ++c.pos;

    // Reject "Sxxx" / "Txxx" - the type letter must stand alone.
    if (isIdentCont(c.peek())) {
        err = "Type annotation must be a single uppercase letter (S or T)";
        errPos = typePos;
        return false;
    }
    return true;
}

// Forward declaration; definition lives further down in this same
// anonymous namespace.
std::string wrapReturnsForExprTk(const std::string& body);

// Parse "function" through the body (exclusive of trailing 'end').
// On entry the cursor is at the 'function' keyword. On success it
// stops one byte past the closing 'end'.
bool parseOneFunction(Cursor& c, EcmdParser::FunctionDef& out,
                      std::string& err, std::size_t& errPos)
{
    out.sourcePos = c.pos;
    const std::string kw = readIdentifier(c);
    if (kw != "function") {
        err = "Expected 'function'";
        errPos = out.sourcePos;
        return false;
    }

    std::size_t commentErr = 0;
    // Local helper: skip trivia and bail with a clear diagnostic on an
    // unterminated comment. Keeps the parse flow readable without
    // hand-repeating the same 5-line check at every separator.
    auto skipOrFail = [&]() -> bool {
        if (skipTrivia(c, commentErr)) return true;
        err = "Unterminated block comment";
        errPos = commentErr;
        return false;
    };

    if (!skipOrFail()) return false;

    if (!isIdentStart(c.peek())) {
        err = "Expected function name after 'function'";
        errPos = c.pos;
        return false;
    }
    out.name = readIdentifier(c);

    if (!skipOrFail()) return false;

    if (c.peek() != '(') {
        err = "Expected '(' after function name";
        errPos = c.pos;
        return false;
    }
    ++c.pos;
    if (!skipOrFail()) return false;

    if (c.peek() != ')') {
        for (;;) {
            EcmdParser::ParamDef p;
            if (!parseOneParam(c, p, err, errPos)) return false;
            for (const auto& existing : out.params) {
                if (existing.name == p.name) {
                    err = "Duplicate parameter name: " + p.name;
                    errPos = p.sourcePos;
                    return false;
                }
            }
            out.params.push_back(std::move(p));

            if (!skipOrFail()) return false;
            if (c.peek() == ',') {
                ++c.pos;
                if (!skipOrFail()) return false;
                continue;
            }
            if (c.peek() == ')') break;

            err = "Expected ',' or ')' in parameter list";
            errPos = c.pos;
            return false;
        }
    }
    ++c.pos;   // consume ')'

    if (!skipOrFail()) return false;

    if (c.peek() == ':') {
        ++c.pos;
        if (!skipOrFail()) return false;
        const std::size_t typePos = c.pos;
        const char t = c.peek();
        if (t == 'T')      out.returnType = EcmdParser::ValueType::Scalar;
        else if (t == 'S') out.returnType = EcmdParser::ValueType::String;
        else {
            if (t == 's' || t == 't') {
                err = "Return type must be uppercase (S or T), got '";
                err += t;
                err += '\'';
            } else {
                err = "Unknown return type '";
                err += (t ? t : '?');
                err += "' - expected S or T";
            }
            errPos = typePos;
            return false;
        }
        ++c.pos;
        if (isIdentCont(c.peek())) {
            err = "Return type must be a single uppercase letter (S or T)";
            errPos = typePos;
            return false;
        }
        if (!skipOrFail()) return false;
    } else {
        out.returnType = EcmdParser::ValueType::Scalar;
    }

    const std::size_t bodyStart = c.pos;
    const std::size_t endPos = findFunctionEnd(c.src, bodyStart, err, errPos);
    if (endPos == kNotFound) return false;

    // Normalise 'return x;' to 'return [x];' so the user can write the
    // intuitive form. ExprTk requires the brackets.
    std::string rawBody(c.src, bodyStart, endPos - bodyStart);
    out.body = wrapReturnsForExprTk(rawBody);
    c.pos = endPos + 3;   // skip past "end"
    return true;
}

// Look ahead from `pos` through trivia (whitespace, comments) and
// report the offset of the first non-trivia byte, or body.size() at EOF.
// Used to decide whether what comes after `return` is already a '['.
std::size_t skipTriviaForLookahead(const std::string& body, std::size_t pos) noexcept {
    while (pos < body.size()) {
        const char c = body[pos];
        if (isAsciiSpace(c)) { ++pos; continue; }
        if (c == '/' && pos + 1 < body.size() && body[pos + 1] == '/') {
            pos += 2;
            while (pos < body.size() && body[pos] != '\n') ++pos;
            continue;
        }
        if (c == '/' && pos + 1 < body.size() && body[pos + 1] == '*') {
            pos += 2;
            while (pos + 1 < body.size()
                && !(body[pos] == '*' && body[pos + 1] == '/')) ++pos;
            if (pos + 1 < body.size()) pos += 2;
            continue;
        }
        break;
    }
    return pos;
}

// Scan a single value expression starting at the cursor and stop at the
// first top-level ';' or '}' (which terminate a return statement), or at
// EOF. The cursor advances to that terminator but does not consume it.
// Strings, comments and nested ()/{}/[] are respected. EOF without a
// terminator is fine - the caller wraps whatever was scanned.
void scanReturnExpression(Cursor& c) noexcept {
    int parenDepth = 0;
    int braceDepth = 0;
    int bracketDepth = 0;
    std::size_t dummy = 0;
    while (!c.atEnd()) {
        const char ec = c.peek();
        if (ec == '\'') { skipStringLiteral(c, dummy); continue; }
        if (ec == '/' && c.peek(1) == '/') { skipLineComment(c); continue; }
        if (ec == '/' && c.peek(1) == '*') { skipBlockComment(c, dummy); continue; }
        if (ec == '(') { ++parenDepth;   ++c.pos; continue; }
        if (ec == ')') { if (parenDepth   > 0) --parenDepth;   ++c.pos; continue; }
        if (ec == '[') { ++bracketDepth; ++c.pos; continue; }
        if (ec == ']') { if (bracketDepth > 0) --bracketDepth; ++c.pos; continue; }
        if (ec == '{') { ++braceDepth;   ++c.pos; continue; }
        if (ec == '}') {
            // An unbalanced '}' is the end of the enclosing block and so
            // also the end of the return expression.
            if (braceDepth > 0) { --braceDepth; ++c.pos; continue; }
            break;
        }
        if (ec == ';' && parenDepth == 0 && braceDepth == 0 && bracketDepth == 0) break;
        ++c.pos;
    }
}

// Wrap the value expression after each `return` keyword in '[...]' so
// users can write 'return x;' the way every other language allows.
// ExprTk itself requires 'return [x];' - this transform makes the
// bracket syntax an internal detail rather than something user-facing.
//
// A return that already has '[' as its next non-trivia token is left
// alone, so users who know ExprTk's native form keep working.
std::string wrapReturnsForExprTk(const std::string& body) {
    std::string out;
    out.reserve(body.size() + 16);

    Cursor c{body, 0};
    std::size_t dummy = 0;

    while (!c.atEnd()) {
        const std::size_t startPos = c.pos;
        const char ch = c.peek();

        // Comments and strings: skip via the shared helpers, then copy
        // the consumed range to the output verbatim so positions inside
        // the original body line up roughly with positions in the
        // rewritten body (helpful when ExprTk reports a parser error).
        if (ch == '/' && c.peek(1) == '/') {
            skipLineComment(c);
            out.append(body, startPos, c.pos - startPos);
            continue;
        }
        if (ch == '/' && c.peek(1) == '*') {
            skipBlockComment(c, dummy);
            out.append(body, startPos, c.pos - startPos);
            continue;
        }
        if (ch == '\'') {
            skipStringLiteral(c, dummy);
            out.append(body, startPos, c.pos - startPos);
            continue;
        }

        if (isIdentStart(ch)) {
            const std::string ident = readIdentifier(c);
            out.append(ident);
            if (ident != "return") continue;

            // After 'return': if the next non-trivia byte is already '[',
            // the user wrote the native ExprTk form - leave it alone.
            const std::size_t afterTrivia = skipTriviaForLookahead(body, c.pos);
            if (afterTrivia < body.size() && body[afterTrivia] == '[') continue;

            // Copy trivia between 'return' and the expression verbatim,
            // then wrap the expression in brackets.
            out.append(body, c.pos, afterTrivia - c.pos);
            c.pos = afterTrivia;

            const std::size_t exprStart = c.pos;
            scanReturnExpression(c);
            out.push_back('[');
            out.append(body, exprStart, c.pos - exprStart);
            out.push_back(']');
            continue;
        }

        // Ordinary byte: pass through.
        out.push_back(ch);
        ++c.pos;
    }

    return out;
}

} // anonymous namespace


EcmdParser::ParseResult EcmdParser::parse(const std::string& fileContent)
{
    ParseResult out;
    Cursor c{fileContent, 0};

    for (;;) {
        std::size_t commentErr = 0;
        if (!skipTrivia(c, commentErr)) {
            out.success = false;
            out.errorMessage = "Unterminated block comment";
            out.errorPos = commentErr;
            out.functions.clear();
            return out;
        }
        if (c.atEnd()) break;

        FunctionDef fn;
        if (!parseOneFunction(c, fn, out.errorMessage, out.errorPos)) {
            out.success = false;
            out.functions.clear();
            return out;
        }

        for (const auto& existing : out.functions) {
            if (existing.name == fn.name) {
                out.success = false;
                out.errorMessage = "Duplicate function name: " + fn.name;
                out.errorPos = fn.sourcePos;
                out.functions.clear();
                return out;
            }
        }
        out.functions.push_back(std::move(fn));
    }

    return out;
}

} // namespace MultiReplaceEngine
