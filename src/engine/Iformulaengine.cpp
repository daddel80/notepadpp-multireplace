// This file is part of MultiReplace.
//
// MultiReplace is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 2 of the License, or
// (at your option) any later version.
//
// IFormulaEngine.cpp
// Definitions for the shared recoverable-skip helpers declared in
// IFormulaEngine.h. Kept out of the header so the LanguageManager and
// Encoding dependencies don't ripple into every translation unit that
// includes the interface.

#include "IFormulaEngine.h"

#include "../LanguageManager.h"
#include "../Encoding.h"

namespace MultiReplaceEngine {

    std::wstring IFormulaEngine::localiseCount(const std::wstring& key,
        std::size_t count)
    {
        return LanguageManager::instance().get(
            key, { std::to_wstring(count) });
    }

    ILuaEngineHost::RecoverableErrorChoice
        IFormulaEngine::handleRecoverableSkip(
            ILuaEngineHost* host,
            const std::wstring& engineName,
            const std::wstring& detailKey,
            const std::string& exprText)
    {
        // Always count the skip so the end-of-run summary stays accurate
        // whether or not a dialog actually surfaces.
        ++_errorSkipCount;

        // Already silenced for this run, or dialogs disabled globally:
        // return SkipOne so the caller treats it as a regular skip.
        if (_skipAllErrors) {
            return ILuaEngineHost::RecoverableErrorChoice::SkipOne;
        }
        if (!host || !host->isFormulaErrorDialogEnabled()) {
            return ILuaEngineHost::RecoverableErrorChoice::SkipOne;
        }

        const std::wstring detailW = LanguageManager::instance().get(
            detailKey, { Encoding::utf8ToWString(exprText) });
        const std::string details = Encoding::wstringToUtf8(detailW);
        const std::string engine = Encoding::wstringToUtf8(engineName);

        const auto choice = host->showRecoverableErrorDialog(engine, details);
        if (choice == ILuaEngineHost::RecoverableErrorChoice::SkipAll) {
            _skipAllErrors = true;
        }
        else if (choice == ILuaEngineHost::RecoverableErrorChoice::Stop) {
            _wantStop = true;
        }
        return choice;
    }

} // namespace MultiReplaceEngine