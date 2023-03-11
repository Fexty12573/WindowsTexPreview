# WindowsTexPreview
Allows windows explorer to display thumbnails for MHW .tex files.
Uses DirectXTex to convert MHW .tex files to Bitmaps which the windows explorer can display.

# Install
(From and admin shell)
```
> regsvr32 path/to/WindowsTexPreview.dll
```

# Uninstall
(From an admin shell)
```
> regsvr32 /u path/to/WindowsTexPreview.dll
```

# Build
Install DirectXTex:
```
> vcpkg install directxtex:x64-windows-static
```

Then build with Visual Studio.
