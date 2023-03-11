#pragma once

#include <thumbcache.h>
#include <wrl.h>

#include <gdiplus.h>
#include <gdiplusheaders.h>

#include <memory>
#include <string>
#include <vector>

#pragma comment(lib, "Gdiplus.lib")

extern const CLSID CLSID_WindowsTexPreview;
extern const wchar_t* CLSID_WindowsTexPreview_STR;

#if 0
#define LogDbgMessage(str) OutputDebugString(str)
#else
#define LogDbgMessage(str)
#endif

#if 0
#define LogMessage(str) OutputDebugString(str)
#else
#define LogMessage(str)
#endif

std::wstring FormatWindowsError(DWORD error);

enum class TexFormat : uint32_t {
    UNKNOWN = 0,
    R8G8B8A8_UNORM = 7,
    R8G8B8A8_UNORM_SRGB = 9,
    R8G8_UNORM = 19,
    BC1_UNORM = 22,
    BC1_UNORM_SRGB = 23,
    BC4_UNORM = 24,
    BC5_UNORM = 26,
    BC6H_UF16 = 28,
    BC7_UNORM = 30,
    BC7_UNORM_SRGB = 31
};

#pragma pack(push, 1)
struct TexHeader {
    char Magic[4];
    int64_t Version;
    int32_t DataBlock;
    int32_t Type;
    int32_t MipCount;
    uint32_t Width;
    uint32_t Height;
    int32_t MipListCount;
    TexFormat Format;
};
#pragma pack(pop)

class CWindowsTexPreview final : public IInitializeWithStream, public IThumbnailProvider {
public:
    CWindowsTexPreview();
    ~CWindowsTexPreview();

    HRESULT Initialize(IStream* pstream, DWORD grfMode) override;
    HRESULT GetThumbnail(UINT cx, HBITMAP* phbmp, WTS_ALPHATYPE* pdwAlpha) override;

    HRESULT QueryInterface(const IID& riid, void** ppvObject) override;
    ULONG AddRef() override;
    ULONG Release() override;

    static UINT ActiveInstances;

private:
    ULONG m_RefCount;
    bool m_Initialized = false;

    Microsoft::WRL::ComPtr<IStream> m_Stream;
};


class CWindowsTexPreviewClassFactory final : public IClassFactory {
public:
    CWindowsTexPreviewClassFactory();
    ~CWindowsTexPreviewClassFactory();

    HRESULT CreateInstance(IUnknown* pUnkOuter, const IID& riid, void** ppvObject) override;
    HRESULT LockServer(BOOL fLock) override;

    HRESULT QueryInterface(const IID& riid, void** ppvObject) override;
    ULONG AddRef() override;
    ULONG Release() override;

private:
    ULONG m_RefCount = 0;
};
