# 4d-plugin-msoffice
Private Research Project

# libxlnt

need to relax errors (bad name, unexpected attribute, bad cell reference...)

but it is still not possible to use this library to process `.xlsm`.

# libopc

need to recpompile [libopc](https://github.com/freuter/libopc) for Visual Studio 2022

# libopenxlsx

latest version (alternative zip) crashed on windows

using snapshot: https://github.com/troldal/OpenXLSX/tree/3eb9c748e3ecd865203fb9946ea86d3c02b3f7d9

> at least the crash on `XLDocument.open()` is avoided, but there is a debug assertion (front() on empty string). ok on release so ignore for now.

# the secret source 

the most proprietary part of this solution has to do with what to copy from the original file without breaking the altered copy

