#include <Windows.h>
#include <spellcheck.h>
#include <string>
#include <iostream>
#include <vector>

// Function to convert wide string to a narrow string
std::string WStringToString(const std::wstring& wstr) {
    if (wstr.empty()) {
        return std::string();
    }
    int count = WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, NULL, 0, NULL, NULL);
    std::string str(count, 0);
    WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, &str[0], count, NULL, NULL);
    return str;
}

// Function to convert narrow string to a wide string
std::wstring StringToWString(const std::string& str) {
    if (str.empty()) {
        return std::wstring();
    }
    int count = MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, NULL, 0);
    std::wstring wstr(count, 0);
    MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, &wstr[0], count);
    return wstr;
}

int main(int argc, char* argv[]) {
    if (argc < 2) {
        std::cout << "Usage: SpellChecker <text_to_check>\n";
        return 1;
    }

    std::wstring text = StringToWString(argv[1]);

    HRESULT hr = CoInitialize(NULL);
    if (SUCCEEDED(hr)) {
        ISpellCheckerFactory* pSpellCheckerFactory = nullptr;
        hr = CoCreateInstance(__uuidof(SpellCheckerFactory), nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pSpellCheckerFactory));
        if (FAILED(hr)) {
            std::cout << "Failed to initialize spell checker factory.\n";
            CoUninitialize();
            return 1;
        }

        ISpellChecker* pSpellChecker = nullptr;
        hr = pSpellCheckerFactory->CreateSpellChecker(L"en-US", &pSpellChecker);
        if (FAILED(hr)) {
            std::cout << "Failed to create spell checker.\n";
            pSpellCheckerFactory->Release();
            CoUninitialize();
            return 1;
        }

        IEnumSpellingError* pEnumErrors = nullptr;
        hr = pSpellChecker->Check(text.c_str(), &pEnumErrors);
        if (FAILED(hr)) {
            std::cout << "Failed to check spelling.\n";
            pSpellChecker->Release();
            pSpellCheckerFactory->Release();
            CoUninitialize();
            return 1;
        }

        std::string jsonOutput = "{ \"Corrections\": [";
        bool firstCorrection = true;

        ISpellingError* pError = nullptr;
        while (pEnumErrors->Next(&pError) == S_OK && pError != nullptr) {
            CORRECTIVE_ACTION action;
            hr = pError->get_CorrectiveAction(&action);

            if (SUCCEEDED(hr)) {
                if (action == CORRECTIVE_ACTION_GET_SUGGESTIONS) {
                    ULONG startIndex, errorLength;
                    pError->get_StartIndex(&startIndex);
                    pError->get_Length(&errorLength);
                    ULONG endIndex = startIndex + errorLength;

                    LPWSTR misspelledWord = new WCHAR[errorLength + 1];
                    wcsncpy_s(misspelledWord, errorLength + 1, text.c_str() + startIndex, errorLength);
                    misspelledWord[errorLength] = L'\0';

                    IEnumString* pSuggestions = nullptr;
                    pSpellChecker->Suggest(misspelledWord, &pSuggestions);

                    if (pSuggestions) {
                        LPWSTR pSuggestion = nullptr;
                        std::string strStart = (firstCorrection ? "" : ", ");
                        std::string suggestionsJson = strStart + "{ \"word\": \"" + WStringToString(misspelledWord)
                            + "\", \"start\": " + std::to_string(startIndex) + ", \"end\": " + std::to_string(endIndex) + ", \"suggestions\": [";
                        bool firstSuggestion = true;

                        while (pSuggestions->Next(1, &pSuggestion, NULL) == S_OK) {
                            if (!firstSuggestion) {
                                suggestionsJson += ", ";
                            }
                            suggestionsJson += "\"" + WStringToString(pSuggestion) + "\"";
                            firstSuggestion = false;
                            CoTaskMemFree(pSuggestion);
                        }
                        suggestionsJson += "] }";
                        jsonOutput += suggestionsJson;
                        pSuggestions->Release();
                        firstCorrection = false;
                    }

                    delete[] misspelledWord;
                }
            }

            pError->Release();
            pError = nullptr;
        }

        jsonOutput += "]}";
        std::cout << jsonOutput << std::endl;

        pEnumErrors->Release();
        pSpellChecker->Release();
        pSpellCheckerFactory->Release();
        CoUninitialize();
    }
    return 0;
}
