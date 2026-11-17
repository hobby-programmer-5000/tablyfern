# 제3자 고지

[English](THIRD-PARTY-NOTICES.md)

Tablyfern 원본 코드와 가상 샘플에는 [MIT](LICENSE)를 적용합니다. 재배포 시 저작권 및 라이선스 고지를 유지하세요. 함께 배포하는 구성 요소는 각자의 조건을 따릅니다.

| 구성 요소 | 릴리스 조건과 역할 |
| --- | --- |
| CPython | PSF License v2와 포함된 구성 요소의 고지를 따릅니다. Windows 포터블 빌드에는 인터프리터와 표준 라이브러리가 포함되며 원문 라이선스는 `licenses/installed/`에 있습니다. 정확한 버전은 `BUILD-INFO.json`에서 확인할 수 있습니다. |
| Qt Base, Qt SVG, PySide6 Essentials 및 Shiboken6 6.12.0 | 배포 라이브러리는 LGPL-3.0을 따르며 구성 요소별 예외와 별도 제3자 조건이 적용됩니다. 원문과 출처 목록은 [licenses/upstream](licenses/upstream) 및 [licenses/SOURCES.json](licenses/SOURCES.json)에 있습니다. |
| XlsxWriter 3.2.9 | BSD-2-Clause입니다. 설치된 원문 고지는 `licenses/installed/`에 복사합니다. [XlsxWriter 릴리스](https://pypi.org/project/XlsxWriter/3.2.9/) |
| Lucide 0.468.0 | ISC를 따르며 Feather에서 유래한 부분에는 MIT 조건이 적용됩니다. 아이콘 제작자는 Lucide Contributors와 Cole Bemis입니다. 저작권 및 허가 고지를 유지하면 상업적 사용, 수정, 재배포가 가능합니다. 원본 SVG `chevron-down`, `files`, `folder-open`, `refresh-cw`, `save`, `sheet`를 함께 배포하며 앱에서 색상을 적용합니다. [해당 버전 SVG 소스](https://github.com/lucide-icons/lucide/tree/0.468.0/icons) · [Lucide 고지](assets/lucide/LICENSE) · [Feather 고지](assets/lucide/FEATHER-LICENSE) |
| PyInstaller 6.22.3 | 부트로더 예외가 있는 GPL과 일부 파일의 Apache-2.0을 따릅니다. 부트로더 예외에 따라 앱 자체 라이선스를 유지할 수 있습니다. [PyInstaller 라이선스](https://pyinstaller.org/en/v6.22.3/license.html) |
| CPython 및 Qt의 바이너리 의존성 | Microsoft VC 런타임은 시스템에서 제공하거나 별도로 설치합니다. OpenSSL, libffi, XZ, zlib, mpdecimal 등은 원래 조건을 유지합니다. `BINARY-MANIFEST.json`, CPython 고지 및 `licenses/`의 원문을 참조하세요. |
| openpyxl 3.1.5 | MIT입니다. 별도 형식 검증 스크립트에 사용합니다. [openpyxl 릴리스](https://pypi.org/project/openpyxl/3.1.5/) |

Windows ZIP은 Qt/PySide/Shiboken 라이브러리를 동적으로 연결합니다. 호환 라이브러리로 교체할 수 있도록 `_internal` 구조를 유지하세요. 이용자는 LGPL이 허용하는 범위에서 라이브러리를 수정하고 그 수정을 위해 결합된 프로그램을 디버깅하거나 역공학할 수 있습니다. 바이너리를 재배포할 때는 고지와 해당 소스에 대한 접근을 유지하세요.

## 해당 버전의 소스

- [Qt Base 6.12.0](https://github.com/qt/qtbase/tree/v6.12.0) 및 [Qt SVG 6.12.0](https://github.com/qt/qtsvg/tree/v6.12.0)
- [PySide 및 Shiboken 6.12.0](https://github.com/pyside/pyside-setup/tree/v6.12.0) 및 [빌드 안내](https://doc.qt.io/qtforpython-6/building_from_source/index.html)
- [Qt LGPL 안내](https://www.qt.io/development/open-source-lgpl-obligations)
- [PyInstaller 6.22.3 소스](https://github.com/pyinstaller/pyinstaller/tree/v6.22.3)
- [CPython 소스 릴리스](https://github.com/python/cpython/tags). 정확한 버전은 `BUILD-INFO.json`을 참조하세요.

## UI·UX 참고 자료

- 익숙한 Windows 컨트롤을 사용하고 열기와 저장 동작을 찾기 쉬운 위치에 배치했습니다. [Microsoft Fluent 2 디자인 원칙](https://fluent2.microsoft.design/design-principles) · [Windows 명령 모음](https://learn.microsoft.com/en-us/windows/apps/design/controls/command-bar)
- 일정한 간격과 글자 크기의 위계를 적용하고 강조색 사용을 줄였습니다. [Windows 배치와 간격](https://learn.microsoft.com/en-us/windows/apps/design/basics/content-basics) · [Windows 색상 지침](https://learn.microsoft.com/en-us/windows/apps/design/signature-experiences/color)
- 중립색 작업 영역과 대비가 뚜렷한 주요 동작은 Material의 배경·텍스트·강조색 역할 구분을 참고했습니다. Qt 화면은 Tablyfern에서 직접 구현한 코드이며 글꼴은 사용자 Windows에 설치된 글꼴을 사용합니다. [Google Material 3 테마 지침](https://developer.android.com/codelabs/m3-design-theming)
- 미리보기에 반영되지 않은 수정, 백그라운드 작업, 오류 상태를 표시합니다. [Nielsen Norman Group 사용성 원칙](https://www.nngroup.com/articles/ten-usability-heuristics/)
- 텍스트 대비와 키보드 포커스의 가시성을 확인했습니다. [W3C 텍스트 대비 지침](https://www.w3.org/WAI/WCAG22/Understanding/contrast-minimum.html) · [W3C 키보드 포커스 지침](https://www.w3.org/WAI/WCAG22/Understanding/focus-visible.html)
