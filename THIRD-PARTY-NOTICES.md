# Third-party notices

[한국어](THIRD-PARTY-NOTICES.ko.md)

Tablyfern's original code and synthetic samples use [MIT](LICENSE). Preserve the copyright and license notice when redistributing those materials. Bundled components keep their respective terms.

| Component | Release terms and role |
| --- | --- |
| CPython | PSF License v2 and incorporated notices. The Windows portable build includes the interpreter, standard library and its full license in `licenses/installed/`. See `BUILD-INFO.json` for the exact interpreter version. |
| Qt Base, Qt SVG, PySide6 Essentials and Shiboken6 6.12.0 | LGPL-3.0 for the distributed libraries, with upstream exceptions and separate third-party terms where applicable. The original terms and attribution inventory are in [licenses/upstream](licenses/upstream) and [licenses/SOURCES.json](licenses/SOURCES.json). |
| XlsxWriter 3.2.9 | BSD-2-Clause. The installed original notice is copied into `licenses/installed/`. [XlsxWriter release](https://pypi.org/project/XlsxWriter/3.2.9/) |
| Lucide 0.468.0 | ISC, with MIT terms for portions derived from Feather. The icons are by Lucide Contributors and Cole Bemis. Commercial use, modification and redistribution are permitted with the applicable copyright and permission notices. The original `chevron-down`, `files`, `folder-open`, `refresh-cw`, `save` and `sheet` SVGs are shipped locally and colored by the app. [Pinned SVG source](https://github.com/lucide-icons/lucide/tree/0.468.0/icons) · [Lucide notice](assets/lucide/LICENSE) · [Feather notice](assets/lucide/FEATHER-LICENSE) |
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

## Design references

- Familiar Windows controls and clear placement of import and save actions. [Microsoft Fluent 2 design principles](https://fluent2.microsoft.design/design-principles) · [Windows command bars](https://learn.microsoft.com/en-us/windows/apps/design/controls/command-bar)
- Consistent spacing, a readable type hierarchy and restrained accent color. [Windows layout and spacing](https://learn.microsoft.com/en-us/windows/apps/design/basics/content-basics) · [Windows color guidance](https://learn.microsoft.com/en-us/windows/apps/design/signature-experiences/color)
- Neutral work surfaces and contrasting primary actions follow Material's separation of surface, text and accent roles. The Qt implementation is original Tablyfern code. Fonts resolve from the user's installed Windows fonts. [Google Material 3 theming guidance](https://developer.android.com/codelabs/m3-design-theming)
- Feedback for unreviewed edits, background work and errors. [Nielsen Norman Group usability heuristics](https://www.nngroup.com/articles/ten-usability-heuristics/)
- Text contrast and visible keyboard focus. [W3C contrast guidance](https://www.w3.org/WAI/WCAG22/Understanding/contrast-minimum.html) · [W3C keyboard focus guidance](https://www.w3.org/WAI/WCAG22/Understanding/focus-visible.html)
