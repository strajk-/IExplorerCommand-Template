// dllmain.cpp : Defines the entry point for the DLL application.
#include "pch.h"
#include <wrl/module.h>
#include <wrl/implements.h>
#include <shobjidl_core.h>
#include <wil/resource.h>
#include <Shellapi.h>
#include <Strsafe.h>
#include <pathcch.h>
#include <vector>
#include "resource.h"

using namespace Microsoft::WRL;

HMODULE g_hModule = nullptr;

BOOL APIENTRY DllMain(HMODULE hModule,
    DWORD ul_reason_for_call,
    LPVOID lpReserved)
{
    if (ul_reason_for_call == DLL_PROCESS_ATTACH)
        g_hModule = hModule;
    return TRUE;
}

// ---------------- Helper Functions -----------------
wil::unique_cotaskmem_string GetExecutableIcon(const wchar_t* exeName)
{
    wil::unique_cotaskmem_string result;
    WCHAR exePath[MAX_PATH] = { 0 };
    bool foundExe = false;

    // --- Check if exeName is already a full path ---
    if (wcschr(exeName, L'\\') || wcschr(exeName, L'/'))
    {
        // Contains path separators, treat as full path
        if (GetFileAttributesW(exeName) != INVALID_FILE_ATTRIBUTES)
        {
            StringCchCopyW(exePath, ARRAYSIZE(exePath), exeName);
            foundExe = true;
        }
    }
    else
    {
        // --- Try PATH environment variable ---
        WCHAR exeNameWithExt[MAX_PATH];
        StringCchCopyW(exeNameWithExt, ARRAYSIZE(exeNameWithExt), exeName);
        if (wcsstr(exeNameWithExt, L".exe") == nullptr)
            StringCchCatW(exeNameWithExt, ARRAYSIZE(exeNameWithExt), L".exe");

        // SearchPath searches current directory, then PATH
        DWORD pathLen = SearchPathW(
            nullptr,              // Use PATH environment variable
            exeNameWithExt,       // File name to search for
            nullptr,              // Extension (already included)
            ARRAYSIZE(exePath),   // Size of output buffer
            exePath,              // Output buffer
            nullptr               // Pointer to filename part (not needed)
        );

        if (pathLen > 0 && pathLen < ARRAYSIZE(exePath))
        {
            foundExe = true;
        }
    }

    // --- If executable found, create icon reference ---
    if (foundExe)
    {
        WCHAR iconRef[MAX_PATH + 8];
        StringCchPrintfW(iconRef, ARRAYSIZE(iconRef), L"%s,%d", exePath, 0);
        result = wil::make_cotaskmem_string_nothrow(iconRef);
        if (result) return result;
    }

    // --- Fallback: DLL embedded icon ---
    WCHAR modulePath[MAX_PATH];
    if (!GetModuleFileNameW(g_hModule, modulePath, ARRAYSIZE(modulePath)))
        return nullptr;

    WCHAR iconReference[MAX_PATH + 16];
    StringCchPrintfW(iconReference, ARRAYSIZE(iconReference), L"%s,-%d",
        modulePath, IDI_ICON_DEFAULT);
    result = wil::make_cotaskmem_string_nothrow(iconReference);
    return result;
}

// Helper method to get the first selected file path
HRESULT GetFirstSelectedFilePath(IShellItemArray* selection, wil::unique_cotaskmem_string& outPath) noexcept try
{
    if (!selection) return E_INVALIDARG;

    DWORD count = 0;
    RETURN_IF_FAILED(selection->GetCount(&count));
    if (count == 0) return E_INVALIDARG;

    ComPtr<IShellItem> item;
    RETURN_IF_FAILED(selection->GetItemAt(0, &item));

    PWSTR filePath = nullptr;
    RETURN_IF_FAILED(item->GetDisplayName(SIGDN_FILESYSPATH, &filePath));
    outPath.reset(filePath);

    return S_OK;
}
catch (...)
{
    return E_FAIL;
}

// Helper method to launch an application with formatted parameters
HRESULT LaunchApplication(PCWSTR exePath, PCWSTR paramFormat, PCWSTR filePath) noexcept try
{
    WCHAR params[4096];
    RETURN_IF_FAILED(StringCchPrintfW(params, ARRAYSIZE(params), paramFormat, filePath));

    SHELLEXECUTEINFOW sei = { sizeof(sei) };
    sei.fMask = SEE_MASK_DEFAULT;
    sei.lpVerb = L"open";
    sei.lpFile = exePath;
    sei.lpParameters = params;
    sei.nShow = SW_SHOWNORMAL;

    if (!ShellExecuteExW(&sei))
        return HRESULT_FROM_WIN32(GetLastError());

    return S_OK;
}
catch (...)
{
    return E_FAIL;
}


// -------------------- Commands --------------------
class MyCommand1 : public RuntimeClass<RuntimeClassFlags<ClassicCom>, IExplorerCommand, IObjectWithSite>
{
public:
    IFACEMETHODIMP GetTitle(IShellItemArray*, PWSTR* name)
    {
        *name = wil::make_cotaskmem_string(L"Edit with Notepad").release();
        return S_OK;
    }

    IFACEMETHODIMP GetIcon(IShellItemArray*, PWSTR* iconPath)
    {
        *iconPath = GetExecutableIcon(L"notepad.exe").release();
        return *iconPath ? S_OK : E_OUTOFMEMORY;
    }

    IFACEMETHODIMP Invoke(IShellItemArray* selection, IBindCtx*) noexcept try
    {
        wil::unique_cotaskmem_string selectedPath;
        RETURN_IF_FAILED(GetFirstSelectedFilePath(selection, selectedPath));

        return LaunchApplication(L"notepad.exe", L"\"%s\"", selectedPath.get());
    }
    catch (...)
    {
        return E_FAIL;
    }

    IFACEMETHODIMP GetToolTip(IShellItemArray*, PWSTR* infoTip) { *infoTip = nullptr; return E_NOTIMPL; }
    IFACEMETHODIMP GetCanonicalName(GUID* guidCommandName) { *guidCommandName = GUID_NULL; return S_OK; }
    IFACEMETHODIMP GetState(IShellItemArray*, BOOL, EXPCMDSTATE* cmdState) { *cmdState = ECS_ENABLED; return S_OK; }
    IFACEMETHODIMP GetFlags(EXPCMDFLAGS* flags) { *flags = ECF_DEFAULT; return S_OK; }
    IFACEMETHODIMP EnumSubCommands(IEnumExplorerCommand** enumCommands) { *enumCommands = nullptr; return E_NOTIMPL; }
    IFACEMETHODIMP SetSite(IUnknown* site) { m_site = site; return S_OK; }
    IFACEMETHODIMP GetSite(REFIID riid, void** site) { return m_site.CopyTo(riid, site); }

protected:
    ComPtr<IUnknown> m_site;
};

class MyCommand2 : public RuntimeClass<RuntimeClassFlags<ClassicCom>, IExplorerCommand, IObjectWithSite>
{
public:
    IFACEMETHODIMP GetTitle(IShellItemArray*, PWSTR* name)
    {
        *name = wil::make_cotaskmem_string(L"Edit with MS Paint").release();
        return S_OK;
    }

    IFACEMETHODIMP GetIcon(IShellItemArray*, PWSTR* iconPath)
    {
        *iconPath = GetExecutableIcon(L"mspaint.exe").release();
        return *iconPath ? S_OK : E_OUTOFMEMORY;
    }

    IFACEMETHODIMP Invoke(IShellItemArray* selection, IBindCtx*) noexcept try
    {
        wil::unique_cotaskmem_string selectedPath;
        RETURN_IF_FAILED(GetFirstSelectedFilePath(selection, selectedPath));

        LaunchApplication(L"mspaint.exe", L"\"%s\"", selectedPath.get());
    }
    catch (...)
    {
        return E_FAIL;
    }

    IFACEMETHODIMP GetToolTip(IShellItemArray*, PWSTR* infoTip) { *infoTip = nullptr; return E_NOTIMPL; }
    IFACEMETHODIMP GetCanonicalName(GUID* guidCommandName) { *guidCommandName = GUID_NULL; return S_OK; }
    IFACEMETHODIMP GetState(IShellItemArray*, BOOL, EXPCMDSTATE* cmdState) { *cmdState = ECS_ENABLED; return S_OK; }
    IFACEMETHODIMP GetFlags(EXPCMDFLAGS* flags) { *flags = ECF_DEFAULT; return S_OK; }
    IFACEMETHODIMP EnumSubCommands(IEnumExplorerCommand** enumCommands) { *enumCommands = nullptr; return E_NOTIMPL; }
    IFACEMETHODIMP SetSite(IUnknown* site) { m_site = site; return S_OK; }
    IFACEMETHODIMP GetSite(REFIID riid, void** site) { return m_site.CopyTo(riid, site); }

protected:
    ComPtr<IUnknown> m_site;
};


// -------------------- Registration --------------------
class __declspec(uuid("E3B9F2A4-5C1D-4E7A-9F0B-8D2C1A7E5F9B")) Command1 final : public MyCommand1 {};
class __declspec(uuid("E3B9F2A4-5C1D-4E7A-9F0B-8D2C1A7E5F8B")) Command2 final : public MyCommand2 {};

CoCreatableClass(Command1)
CoCreatableClass(Command2)

STDAPI DllGetActivationFactory(_In_ HSTRING activatableClassId, _COM_Outptr_ IActivationFactory** factory)
{
    return Module<ModuleType::InProc>::GetModule().GetActivationFactory(activatableClassId, factory);
}

STDAPI DllCanUnloadNow()
{
    return Module<InProc>::GetModule().GetObjectCount() == 0 ? S_OK : S_FALSE;
}

STDAPI DllGetClassObject(_In_ REFCLSID rclsid, _In_ REFIID riid, _COM_Outptr_ void** instance)
{
    return Module<InProc>::GetModule().GetClassObject(rclsid, riid, instance);
}