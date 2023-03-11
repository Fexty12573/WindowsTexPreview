// dllmain.cpp : Defines the entry point for the DLL application.
#include "pch.h"
#include <Shlwapi.h>
#include <ShlObj.h>

#include "WindowsTexPreview.h"
#include "Reg.h"

#pragma comment(lib, "Shlwapi.lib")

static HMODULE g_Module = nullptr;
LONG g_DllRefCount = 0;
ULONG_PTR g_GdiplusToken;

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv) {
    LogMessage(L"[WTP] DllGetClassObject\n");
    if (rclsid != CLSID_WindowsTexPreview) {
        return CLASS_E_CLASSNOTAVAILABLE;
    }

    const auto factory = new CWindowsTexPreviewClassFactory();
    if (factory == nullptr) {
        return E_OUTOFMEMORY;
    }

    const auto hr = factory->QueryInterface(riid, ppv);
    factory->Release();
    return hr;
}

STDAPI DllCanUnloadNow() {
    return g_DllRefCount == 0 ? S_OK : S_FALSE;
}

#define ERROR_IF_FAILED(expr) do{ const auto _res = expr;\
    if (_res != ERROR_SUCCESS) {\
        LogMessage(FormatWindowsError(_res).c_str());\
        return HRESULT_FROM_WIN32(_res);\
    }\
} while (false)

STDAPI DllRegisterServer() {
    LogMessage(L"[WTP] DllRegisterServer\n");
    const auto path = std::make_unique<wchar_t[]>(MAX_PATH);
    if (path == nullptr) {
        return E_OUTOFMEMORY;
    }

    if (GetModuleFileNameW(g_Module, path.get(), MAX_PATH) == 0) {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    ERROR_IF_FAILED(RegisterInprocServer(path.get(), CLSID_WindowsTexPreview, L"WindowsTexPreview", L"Apartment"));
    ERROR_IF_FAILED(RegisterShellExtThumbnailHandler(L".tex", CLSID_WindowsTexPreview));

    SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, nullptr, nullptr);

    return S_OK;
}

STDAPI DllUnregisterServer() {
    LogMessage(L"[WTP] DllUnregisterServer\n");
    const auto path = std::make_unique<wchar_t[]>(MAX_PATH);

    if (path == nullptr) {
        return E_OUTOFMEMORY;
    }

    if (GetModuleFileNameW(g_Module, path.get(), MAX_PATH) == 0) {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    ERROR_IF_FAILED(UnregisterInprocServer(CLSID_WindowsTexPreview));
    ERROR_IF_FAILED(UnregisterShellExtThumbnailHandler(L".tex"));

    return S_OK;
}


BOOL APIENTRY DllMain(HMODULE module, DWORD reason, LPVOID) {
    if (reason == DLL_PROCESS_ATTACH) {
        Gdiplus::GdiplusStartupInput gdiplusStartupInput;
        Gdiplus::GdiplusStartup(&g_GdiplusToken, &gdiplusStartupInput, nullptr);

        LogMessage(L"[WTP] DLL_PROCESS_ATTACH\n");
        g_Module = module;
        DisableThreadLibraryCalls(module);
    } else if (reason == DLL_PROCESS_DETACH) {
        LogMessage(L"[WTP] DLL_PROCESS_DETACH\n");
        Gdiplus::GdiplusShutdown(g_GdiplusToken);
    }

    return TRUE;
}


std::wstring FormatWindowsError(DWORD error) {
    wchar_t* buffer = nullptr;
    const auto size = FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS, nullptr, error, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), reinterpret_cast<LPWSTR>(&buffer), 0, nullptr);
    if (size == 0) {
        return L"";
    }

    auto message = L"[WTP] Error: " + std::wstring(buffer, size);
    LocalFree(buffer);
    return message;
}

