#include "pch.h"
#include "WindowsTexPreview.h"

#include <bitset>
#include <vector>

#include <DirectXTex.h>
#include <Shlwapi.h>

#define MAKEFOURCC_(a, b, c, d) (static_cast<uint32_t>(a) | (static_cast<uint32_t>(b) << 8) | (static_cast<uint32_t>(c) << 16) | (static_cast<uint32_t>(d) << 24))
#define MAKEFOURCC(fourcc_str) MAKEFOURCC_((fourcc_str)[0], (fourcc_str)[1], (fourcc_str)[2], (fourcc_str)[3])


struct DDS_PIXELFORMAT {
    uint32_t    size;
    uint32_t    flags;
    uint32_t    fourCC;
    uint32_t    RGBBitCount;
    uint32_t    RBitMask;
    uint32_t    GBitMask;
    uint32_t    BBitMask;
    uint32_t    ABitMask;
};

struct DDS_HEADER {
    uint32_t        magic; // "DDS "
    uint32_t        size;
    uint32_t        flags;
    uint32_t        height;
    uint32_t        width;
    uint32_t        pitchOrLinearSize;
    uint32_t        depth; // only if DDS_HEADER_FLAGS_VOLUME is set in flags
    uint32_t        mipMapCount;
    uint32_t        reserved1[11]; // 72
    DDS_PIXELFORMAT ddspf; // 104
    uint32_t        caps;
    uint32_t        caps2;
    uint32_t        caps3;
    uint32_t        caps4;
    uint32_t        reserved2;
};

struct DDS_HEADER_DXT10 {
    DXGI_FORMAT     dxgiFormat;
    uint32_t        resourceDimension;
    uint32_t        miscFlag; // see D3D11_RESOURCE_MISC_FLAG
    uint32_t        arraySize;
    uint32_t        miscFlags2;
};

enum DDS_FLAGS : uint32_t {
    DDSD_CAPS = 0x1,
    DDSD_HEIGHT = 0x2,
    DDSD_WIDTH = 0x4,
    DDSD_PITCH = 0x8,
    DDSD_PIXELFORMAT = 0x1000,
    DDSD_MIPMAPCOUNT = 0x20000,
    DDSD_LINEARSIZE = 0x80000,
    DDSD_DEPTH = 0x800000
};

enum DDS_PIXELFORMAT_FLAGS : uint32_t {
    DDPF_ALPHAPIXELS = 0x1,
    DDPF_ALPHA = 0x2,
    DDPF_FOURCC = 0x4,
    DDPF_RGB = 0x40,
    DDPF_YUV = 0x200,
    DDPF_LUMINANCE = 0x20000
};

enum DDS_PIXELDATA_COMPRESSION : uint32_t {
    DDSCAPS_COMPLEX = 0x8,
    DDSCAPS_TEXTURE = 0x1000,
    DDSCAPS_MIPMAP = 0x400000
};

struct DDS_EXTENDED_HEADER {
    DDS_HEADER hdr;
    DDS_HEADER_DXT10 ext;
};

DXGI_FORMAT ConvertFormat(TexFormat fmt) {
    switch (fmt) {
    case TexFormat::UNKNOWN: return DXGI_FORMAT_UNKNOWN;
    case TexFormat::R8G8B8A8_UNORM: return DXGI_FORMAT_R8G8B8A8_UNORM;
    case TexFormat::R8G8B8A8_UNORM_SRGB: return DXGI_FORMAT_R8G8B8A8_UNORM_SRGB;
    case TexFormat::R8G8_UNORM: return DXGI_FORMAT_R8G8_UNORM;
    case TexFormat::BC1_UNORM: return DXGI_FORMAT_BC1_UNORM;
    case TexFormat::BC1_UNORM_SRGB: return DXGI_FORMAT_BC1_UNORM_SRGB;
    case TexFormat::BC4_UNORM: return DXGI_FORMAT_BC4_UNORM;
    case TexFormat::BC5_UNORM: return DXGI_FORMAT_BC5_UNORM;
    case TexFormat::BC6H_UF16: return DXGI_FORMAT_BC6H_UF16;
    case TexFormat::BC7_UNORM: return DXGI_FORMAT_BC7_UNORM;
    case TexFormat::BC7_UNORM_SRGB: return DXGI_FORMAT_BC7_UNORM_SRGB;
    default: break;
    }

    return DXGI_FORMAT_UNKNOWN;
}

uint32_t FormatToFourCC(TexFormat fmt) {
    switch (fmt) {
    case TexFormat::UNKNOWN:             return MAKEFOURCC("UNKN");
    case TexFormat::R8G8B8A8_UNORM: [[fallthrough]];
    case TexFormat::R8G8B8A8_UNORM_SRGB: [[fallthrough]];
    case TexFormat::BC6H_UF16: [[fallthrough]];
    case TexFormat::BC7_UNORM: [[fallthrough]];
    case TexFormat::R8G8_UNORM: [[fallthrough]];
    case TexFormat::BC1_UNORM_SRGB: [[fallthrough]];
    case TexFormat::BC7_UNORM_SRGB:      return MAKEFOURCC("DX10");
    case TexFormat::BC1_UNORM:           return MAKEFOURCC("DXT1");
    case TexFormat::BC4_UNORM:           return MAKEFOURCC("BC4U");
    case TexFormat::BC5_UNORM:           return MAKEFOURCC("BC5U");
    default:                          break;
    }

    return MAKEFOURCC("UNKN");
}

bool FormatIs4bpp(TexFormat fmt) {
    return fmt == TexFormat::BC1_UNORM || fmt == TexFormat::BC1_UNORM_SRGB || fmt == TexFormat::BC4_UNORM;
}

bool FormatIs16bpp(TexFormat fmt) {
    return fmt == TexFormat::R8G8_UNORM;
}


// ===========================================================================
// CWindowsTexPreview
// ===========================================================================

// {B802F092-43DB-434D-9A0A-C1837D858487}
extern "C" const GUID CLSID_WindowsTexPreview = {
    0xb802f092, 0x43db, 0x434d, { 0x9a, 0xa, 0xc1, 0x83, 0x7d, 0x85, 0x84, 0x87 }
};

UINT CWindowsTexPreview::ActiveInstances = 0;

extern LONG g_DllRefCount;

HRESULT CWindowsTexPreview::QueryInterface(const IID& riid, void** ppvObject) {
    LogMessage(L"[WTP] CWindowsTexPreview::QueryInterface");

    static const QITAB qit[] = {
        QITABENT(CWindowsTexPreview, IInitializeWithStream),
        QITABENT(CWindowsTexPreview, IThumbnailProvider),
        { nullptr },
    };

    return QISearch(this, qit, riid, ppvObject);
}

ULONG CWindowsTexPreview::AddRef() {
    return InterlockedIncrement(&m_RefCount);
}

ULONG CWindowsTexPreview::Release() {
    const ULONG count = InterlockedDecrement(&m_RefCount);
    if (count == 0) {
        delete this;
    }

    return count;
}

#define RETURN_IF_FAILED(expr, msg) do {\
    const HRESULT _hr = expr;\
    if (FAILED(_hr)) {\
        LogMessage(msg);\
        LogMessage(FormatWindowsError(_hr).c_str());\
        return _hr;\
    }\
} while (0)

CWindowsTexPreview::CWindowsTexPreview() : m_RefCount(1) {
    LogMessage(L"[WTP] CWindowsTexPreview::CWindowsTexPreview");
    InterlockedIncrement(&g_DllRefCount);
}

CWindowsTexPreview::~CWindowsTexPreview() {
    InterlockedDecrement(&g_DllRefCount);
}

HRESULT CWindowsTexPreview::Initialize(IStream* pstream, DWORD grfMode) {
    LogMessage(L"[WTP] Invoked CWindowsTexPreview::Initialize");

    if (!m_Initialized) {
        m_Stream = pstream;
        m_Initialized = true;

        return S_OK;
    }

    return E_UNEXPECTED;
}

HRESULT CWindowsTexPreview::GetThumbnail(UINT cx, HBITMAP* phbmp, WTS_ALPHATYPE* pdwAlpha) {
    LogMessage((L"[WTP] GetThumbnail: " + std::to_wstring(cx)).c_str());

    TexHeader header{};
    ULONG nread = 0;
    ULARGE_INTEGER newpos = { 0 };

    RETURN_IF_FAILED(m_Stream->Read(&header, sizeof header, &nread), L"[WTP] Failed to read header: ");

    if (header.Format == TexFormat::UNKNOWN) {
        LogMessage(L"[WTP] Unknown format");
        return E_FAIL;
    }

    RETURN_IF_FAILED(m_Stream->Seek({ .QuadPart = 0xB8 }, STREAM_SEEK_SET, &newpos), L"[WTP] Failed to seek to offset 0xB8: ");

    int64_t offset = 0;
    RETURN_IF_FAILED(m_Stream->Read(&offset, sizeof offset, &nread), L"[WTP] Failed to read data offset: ");

    STATSTG stat{};
    RETURN_IF_FAILED(m_Stream->Stat(&stat, STATFLAG_NONAME), L"[WTP] Failed to stat stream: ");

    const uint32_t four_cc = FormatToFourCC(header.Format);
    const uint32_t header_size = four_cc == MAKEFOURCC("DX10") ? sizeof(DDS_EXTENDED_HEADER) : sizeof(DDS_HEADER);

    std::vector<uint8_t> data;
    data.resize(header_size + stat.cbSize.QuadPart - offset);

    RETURN_IF_FAILED(m_Stream->Seek({ .QuadPart = offset }, STREAM_SEEK_SET, &newpos), L"[WTP] Failed to seek to data offset: ");
    RETURN_IF_FAILED(m_Stream->Read(data.data() + header_size, static_cast<ULONG>(data.size()) - header_size, &nread), L"[WTP] Failed to read data: ");

    const auto dds = reinterpret_cast<DDS_EXTENDED_HEADER*>(data.data());

    dds->hdr.magic = MAKEFOURCC("DDS ");
    dds->hdr.size = sizeof DDS_HEADER - 4; // Original DDS_HEADER does not contain the 'magic' field
    dds->hdr.flags = DDSD_CAPS | DDSD_HEIGHT | DDSD_WIDTH | DDSD_PIXELFORMAT | DDSD_MIPMAPCOUNT | DDSD_LINEARSIZE;
    dds->hdr.height = header.Height;
    dds->hdr.width = header.Width;

    if (FormatIs4bpp(header.Format)) {
        dds->hdr.pitchOrLinearSize = header.Width * header.Height / 2;
    } else if (FormatIs16bpp(header.Format)) {
        dds->hdr.pitchOrLinearSize = header.Width * header.Height * 2;
    } else {
        dds->hdr.pitchOrLinearSize = header.Width * header.Height;
    }

    dds->hdr.depth = 1;
    dds->hdr.mipMapCount = header.MipCount;
    dds->hdr.ddspf.size = sizeof DDS_PIXELFORMAT;
    dds->hdr.ddspf.flags = DDPF_FOURCC;
    dds->hdr.ddspf.fourCC = four_cc;
    dds->hdr.caps = DDSCAPS_COMPLEX | DDSCAPS_MIPMAP | DDSCAPS_TEXTURE;

    if (four_cc == MAKEFOURCC("DX10")) {
        dds->ext.dxgiFormat = ConvertFormat(header.Format);
        dds->ext.resourceDimension = D3D11_RESOURCE_DIMENSION_TEXTURE2D;
        dds->ext.arraySize = 1;
    }

    DirectX::TexMetadata metadata{};
    DirectX::ScratchImage scratch_image;
    DirectX::ScratchImage converted;
    RETURN_IF_FAILED(DirectX::LoadFromDDSMemory(data.data(), data.size(), DirectX::DDS_FLAGS_NONE, &metadata, scratch_image), L"Failed to load DDS: ");

    LogDbgMessage(L"[WTP] Converting Image...");

    if (DirectX::IsCompressed(metadata.format)) {
        RETURN_IF_FAILED(DirectX::Decompress(
            scratch_image.GetImages(),
            scratch_image.GetImageCount(),
            scratch_image.GetMetadata(),
            DXGI_FORMAT_B8G8R8A8_UNORM,
            converted
        ), L"Failed to decompress image: ");
    } else {
        RETURN_IF_FAILED(DirectX::Convert(
            scratch_image.GetImages(),
            scratch_image.GetImageCount(),
            scratch_image.GetMetadata(),
            DXGI_FORMAT_B8G8R8A8_UNORM,
            DirectX::TEX_FILTER_DEFAULT,
            DirectX::TEX_THRESHOLD_DEFAULT,
            converted
        ), L"Failed to convert image: ");
    }

    LogDbgMessage(L"[WTP] Resizing Image...");

    const float original_aspect = (float)header.Width / (float)header.Height;
    constexpr float target_aspect = 1.0f;
    const float scaling_factor = original_aspect > target_aspect ? (float)header.Height / (float)cx : (float)header.Width / (float)cx;

    const int new_width = (int)((float)header.Width / scaling_factor);
    const int new_height = (int)((float)header.Height / scaling_factor);

    DirectX::ScratchImage resized;
    RETURN_IF_FAILED(DirectX::Resize(
        converted.GetImages(),
        converted.GetImageCount(),
        converted.GetMetadata(),
        new_width,
        new_height,
        DirectX::TEX_FILTER_DEFAULT,
        resized
    ), L"Failed to resize image: ");

    LogDbgMessage(L"[WTP] Converting to Bitmap...");

    const DirectX::Image* image = resized.GetImage(0, 0, 0);
    if (!image) {
        LogMessage(L"[WTP] Failed to get image");
        return E_FAIL;
    }

    LogDbgMessage(L"[WTP] Creating Bitmap...");

    constexpr Gdiplus::PixelFormat format = PixelFormat32bppARGB;
    Gdiplus::Bitmap bmp(new_width, new_height, new_width * 4, format, image->pixels);

    LogDbgMessage(L"[WTP] Retrieving HBITMAP...");

    if (const auto status = bmp.GetHBITMAP(Gdiplus::Color(0, 0, 0, 0), phbmp)) {
        LogMessage((L"[WTP] Failed to get HBITMAP: " + std::to_wstring(status)).c_str());
        return E_FAIL;
    }

    *pdwAlpha = WTSAT_ARGB;

    LogMessage(L"[WTP] GetThumbnail done");
    return S_OK;
}

// ===========================================================================
// CWindowsTexPreviewClassFactory
// ===========================================================================

CWindowsTexPreviewClassFactory::CWindowsTexPreviewClassFactory() : m_RefCount(1) {
    LogMessage(L"[WTP] CWindowsTexPreviewClassFactory::CWindowsTexPreviewClassFactory");
    InterlockedIncrement(&g_DllRefCount);
}

CWindowsTexPreviewClassFactory::~CWindowsTexPreviewClassFactory() {
    InterlockedDecrement(&g_DllRefCount);
}

HRESULT CWindowsTexPreviewClassFactory::CreateInstance(IUnknown* pUnkOuter, const IID& riid, void** ppvObject) {
    LogMessage(L"[WTP] CWindowsTexPreviewClassFactory::CreateInstance");
    if (pUnkOuter != nullptr) {
        return CLASS_E_NOAGGREGATION;
    }

    const auto object = new CWindowsTexPreview();
    if (object == nullptr) {
        return E_OUTOFMEMORY;
    }

    const auto hr = object->QueryInterface(riid, ppvObject);
    object->Release();
    return hr;
}

HRESULT CWindowsTexPreviewClassFactory::LockServer(BOOL fLock) {
    LogMessage(L"[WTP] CWindowsTexPreviewClassFactory::LockServer");
    if (fLock) {
        InterlockedIncrement(&g_DllRefCount);
    } else {
        InterlockedDecrement(&g_DllRefCount);
    }

    return S_OK;
}

HRESULT CWindowsTexPreviewClassFactory::QueryInterface(const IID& riid, void** ppvObject) {
    LogMessage(L"[WTP] CWindowsTexPreviewClassFactory::QueryInterface");
    static const QITAB qit[] = {
        QITABENT(CWindowsTexPreviewClassFactory, IClassFactory),
        { nullptr },
    };

    return QISearch(this, qit, riid, ppvObject);
}

ULONG CWindowsTexPreviewClassFactory::AddRef() {
    return InterlockedIncrement(&m_RefCount);
}

ULONG CWindowsTexPreviewClassFactory::Release() {
    const ULONG count = InterlockedDecrement(&m_RefCount);
    if (count == 0) {
        delete this;
    }

    return count;
}
