# Third-party notices

[한국어](THIRD-PARTY-NOTICES.ko.md)

Tablyfern's original code and synthetic samples use [MIT](LICENSE). Preserve the copyright and license notice when redistributing those materials. Bundled components keep their respective terms.

| Component | Release terms and role |
| --- | --- |
| CPython | PSF License v2 and incorporated notices. The Windows portable build includes the interpreter, standard library and its full license in `licenses/installed/`. See `BUILD-INFO.json` for the exact interpreter version. |
| Qt Base, Qt SVG, PySide6 Essentials and Shiboken6 6.12.0 | LGPL-3.0 for the distributed libraries, with upstream exceptions and separate third-party terms where applicable. The original terms and attribution inventory are in [licenses/upstream](licenses/upstream) and [licenses/SOURCES.json](licenses/SOURCES.json). |
| XlsxWriter 3.2.9 | BSD-2-Clause. The installed original notice is copied into `licenses/installed/`. [XlsxWriter release](https://pypi.org/project/XlsxWriter/3.2.9/) |
| PyInstaller 6.22.3 | GPL with the bootloader exception and Apache-2.0 on selected files. The bootloader exception allows the application to keep its own license. [PyInstaller license](https://pyinstaller.org/en/v6.22.3/license.html) |
| CPython and Qt binary dependencies | The Microsoft VC runtime is provided by the operating system or installed separately. Incorporated libraries such as OpenSSL, libffi, XZ, zlib and mpdecimal retain their original terms. See `BINARY-MANIFEST.json`, the full CPython notice and the original texts in `licenses/`. |
| openpyxl 3.1.5 | MIT. Used by the separate format verification runner. [openpyxl release](https://pypi.org/project/openpyxl/3.1.5/) |

The Windows ZIP uses dynamically linked Qt/PySide/Shiboken libraries. Keep the `_internal` layout so compatible libraries can be replaced. Recipients may modify the LGPL libraries and debug or reverse-engineer the combined work for those modifications as permitted by LGPL. Preserve these notices and corresponding source access when redistributing the binary.

## Corresponding source

- [Qt Base 6.12.0](https://github.com/qt/qtbase/tree/v6.12.0) and [Qt SVG 6.12.0](https://github.com/qt/qtsvg/tree/v6.12.0)
- [PySide and Shiboken 6.12.0](https://github.com/pyside/pyside-setup/tree/v6.12.0) and [build instructions](https://doc.qt.io/qtforpython-6/building_from_source/index.html)
- [Qt LGPL guidance](https://www.qt.io/development/open-source-lgpl-obligations)
- [PyInstaller 6.22.3 source](https://github.com/pyinstaller/pyinstaller/tree/v6.22.3)
- [CPython source releases](https://github.com/python/cpython/tags), with the exact release identified by `BUILD-INFO.json`
