#include <windows.h>
#include <stddef.h>

#if defined(_M_IX86)
#pragma comment(linker, "/entry:ShortcutDllEntry@12")
#pragma comment(linker, "/export:xlAutoOpen=_xlAutoOpen@0")
#pragma comment(linker, "/export:xlAutoClose=_xlAutoClose@0")
#pragma comment(linker, "/export:ClairReturnFromWorksheet=_ClairReturnFromWorksheet@0")
#pragma comment(linker, "/export:ClairSetShortcutCallback=_ClairSetShortcutCallback@4")
#else
#pragma comment(linker, "/entry:ShortcutDllEntry")
#pragma comment(linker, "/export:xlAutoOpen")
#pragma comment(linker, "/export:xlAutoClose")
#pragma comment(linker, "/export:ClairReturnFromWorksheet")
#pragma comment(linker, "/export:ClairSetShortcutCallback")
#endif

using ShortcutCallback = void (WINAPI*)();

namespace
{
    // Only the XLOPER12 fields used by this command. Both MSVC architectures
    // place the DWORD type after a 24-byte value union.
    struct ExcelValue
    {
        union
        {
            double number;
            wchar_t* text;
            int integer;
            BYTE storage[24];
        } value;
        DWORD type;
    };

    static_assert(offsetof(ExcelValue, type) == 24 && sizeof(ExcelValue) == 32, "XLOPER12 ABI mismatch");
    static_assert(sizeof(wchar_t) == 2, "Excel requires UTF-16");

    constexpr int ExcelSuccess = 0;
    constexpr int ExcelFree = 0x4000;
    constexpr int ExcelGetName = 0x4009;
    constexpr int ExcelRegister = 149;
    constexpr DWORD ExcelNumber = 1;
    constexpr DWORD ExcelText = 2;
    constexpr DWORD ExcelInteger = 0x0800;

    using ExcelCall = int (WINAPI*)(int function, int count, ExcelValue** arguments, ExcelValue* result);

    template<size_t Size>
    struct CountedText
    {
        wchar_t characters[Size];
        constexpr CountedText(const wchar_t (&text)[Size]) : characters{}
        {
            characters[0] = static_cast<wchar_t>(Size - 1);
            for (size_t index = 1; index < Size; ++index)
                characters[index] = text[index - 1];
        }
    };

    CountedText commandName(L"ClairReturnFromWorksheet");
    CountedText commandSignature(L"J");
    CountedText emptyText(L"");
    PVOID volatile shortcutCallback = nullptr;

    ExcelValue Text(wchar_t* text)
    {
        ExcelValue value{};
        value.value.text = text;
        value.type = ExcelText;
        return value;
    }
}

extern "C" BOOL WINAPI ShortcutDllEntry(HINSTANCE, DWORD, LPVOID)
{
    return TRUE;
}

extern "C" void WINAPI ClairSetShortcutCallback(ShortcutCallback callback)
{
    InterlockedExchangePointer(&shortcutCallback, reinterpret_cast<PVOID>(callback));
}

extern "C" int WINAPI ClairReturnFromWorksheet()
{
    auto callback = reinterpret_cast<ShortcutCallback>(
        InterlockedCompareExchangePointer(&shortcutCallback, nullptr, nullptr));
    if (callback) callback();
    return 1;
}

extern "C" int WINAPI xlAutoOpen()
{
    HMODULE host = GetModuleHandleW(nullptr);
    if (!host) return 0;

    // Excel's vector callback takes the result as its last argument.
#pragma warning(suppress:4191)
    auto excel = reinterpret_cast<ExcelCall>(GetProcAddress(host, "MdCallBack12"));
    if (!excel) return 0;

    ExcelValue module{};
    if (excel(ExcelGetName, 0, nullptr, &module) != ExcelSuccess) return 0;

    auto entry = Text(commandName.characters);
    auto signature = Text(commandSignature.characters);
    auto name = Text(commandName.characters);
    auto arguments = Text(emptyText.characters);
    ExcelValue kind{};
    kind.type = ExcelInteger;
    kind.value.integer = 2;
    ExcelValue* inputs[] = { &module, &entry, &signature, &name, &arguments, &kind };
    ExcelValue registration{};
    const int result = excel(ExcelRegister, _countof(inputs), inputs, &registration);

    ExcelValue* release[] = { &module };
    excel(ExcelFree, _countof(release), release, nullptr);
    return result == ExcelSuccess && registration.type == ExcelNumber ? 1 : 0;
}

extern "C" int WINAPI xlAutoClose()
{
    ClairSetShortcutCallback(nullptr);
    return 1;
}
