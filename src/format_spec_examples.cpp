// Generator: prints two paste-friendly tables of FormatSpec
// examples. Run after format_spec_qa to copy real outputs into
// MR-style notes or docs.

#include "exprtk/FormatSpec.h"

#include <iostream>
#include <string>
#include <vector>

namespace {

struct Sample {
    std::wstring spec;
    double value;
    std::wstring note;
};

void runSheet(const std::wstring& title, const std::vector<Sample>& s) {
    std::wcout << L"\n=== " << title << L" ===\n";
    std::wcout << L"spec          value           output          note\n";
    std::wcout << L"------------  --------------  --------------  ----\n";
    for (const auto& smp : s) {
        auto p = FormatSpec::parse(smp.spec);
        std::wstring out = p.valid ? FormatSpec::apply(p, smp.value)
                                   : (L"<error: " + std::wstring(p.errorMessage.begin(), p.errorMessage.end()) + L">");
        std::wprintf(L"%-12ls  %-14g  %-14ls  %ls\n",
                     smp.spec.c_str(), smp.value, out.c_str(), smp.note.c_str());
    }
}

}  // namespace

int main() {

    // -------- Sheet 1: Numeric --------
    runSheet(L"Numeric format examples", {
        // type-only
        {L"f",      3.14159,     L"default precision = 6"},
        {L"e",      12345.0,     L"scientific"},
        {L"g",      0.0001,      L"general - smart"},
        {L"x",      255.0,       L"hex"},
        {L"o",      8.0,         L"octal"},
        {L"b",      5.0,         L"binary"},

        // width + zeropad
        {L"10f",    3.14,        L"width pads with spaces"},
        {L"010f",   3.14,        L"zero pad"},
        {L"08x",    255.0,       L"hex zero pad"},
        {L"8b",     5.0,         L"binary right-aligned"},

        // precision
        {L".2f",    3.14159,     L"2 decimals"},
        {L".0f",    3.7,         L"no decimals, rounds"},
        {L"05.2f",  3.14,        L"05 width, 2 decimals"},

        // force sign
        {L"+f",     3.14,        L"force + on positive"},
        {L"+f",     -3.14,       L"keeps - on negative"},
        {L"+8.2f",  3.14,        L"sign + width + precision"},

        // min-max precision (trailing zero trim)
        {L".2-5f",  3.1,         L"keeps 2 trailing zeros"},
        {L".2-5f",  3.14159,     L"shows all up to 5"},
        {L".2-5f",  3.141592,    L"rounds at 5"},
        {L".0-3f",  3.0,         L"strips trailing dot"},
        {L".0-3f",  3.14,        L"shows 2 decimals"},

        // default type (no letter)
        {L"+",      3.14,        L"shortest, sign forced"},
        {L"08",     3.14,        L"shortest, width 8 zero-pad"},
    });

    // -------- Sheet 2: Duration --------
    runSheet(L"Duration format examples", {
        // seconds input
        {L"ts:ms",   90.0,        L"1m 30s"},
        {L"ts:hms",  3725.0,      L"1h 2m 5s"},
        {L"ts:hms",  86400.0,     L"24 hours"},

        // minutes input
        {L"tm:hm",   90.0,        L"work day 1h30m"},
        {L"tm:hm",   480.0,       L"8h workday"},
        {L"tm:hms",  90.0,        L"with seconds always :00"},

        // hours input
        {L"th:hms",  2.5,         L"2.5 hours"},
        {L"th:dh",   25.0,        L"1 day 1 hour"},

        // days input
        {L"td:dhms", 1.5,         L"1d 12h"},
        {L"td:dhm",  3.25,        L"3d 6h"},
        {L"td:dh",   2.0,         L"2 full days"},

        // negatives
        {L"ts:hms",  -3725.0,     L"negative durations"},
    });

    return 0;
}