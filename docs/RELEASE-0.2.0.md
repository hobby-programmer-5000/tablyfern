# Tablyfern 0.2.0

This Windows x64 release streams large JSON and JSON Lines files to temporary disk storage. It removes Tablyfern's fixed caps on file size, pasted text, record length, batch count and CSV/TSV cell count. It adds `.jsonlines` and `.ldjson` input and detects JSON or JSON Lines in files with other extensions.

The **Limits** dialog and [user guide](https://github.com/hobby-programmer-5000/tablyfern/blob/v0.2.0/README.md#file-and-worksheet-limits) describe the remaining constraints. Each input file maps to one sheet with up to 1,048,575 data rows plus a header and 16,384 columns. Long cell text continues in adjacent columns. Available memory, temporary disk space and workbook package size affect actual capacity.

Extract the entire Windows x64 ZIP and run `Tablyfern.exe`. Use `SHA256SUMS.txt` to verify the download. [Third-party notices and corresponding sources](https://github.com/hobby-programmer-5000/tablyfern/blob/v0.2.0/THIRD-PARTY-NOTICES.md)

## 한국어

이번 Windows x64 버전은 큰 JSON과 JSONL을 읽어 디스크 임시 공간을 사용합니다. 파일 크기, 붙여넣기 크기, 레코드 길이, 묶음 파일 수, CSV·TSV 셀 수에 앱이 정한 고정 한도를 없앴습니다. `.jsonlines`와 `.ldjson`을 지원하며, 다른 확장자도 JSON 또는 JSONL로 자동 감지합니다.

앱의 **사용 한도** 창과 [한글 안내](https://github.com/hobby-programmer-5000/tablyfern/blob/v0.2.0/README.ko.md#파일-및-시트-한도)에서 적용되는 한도를 확인할 수 있습니다. 파일 하나당 시트 한 장에 머리글 외 데이터 1,048,575행과 16,384열까지 저장합니다. 긴 셀 텍스트는 옆 열로 이어집니다. 실제 처리 가능 크기는 메모리, 디스크 임시 공간, 통합 문서 패키지 크기에 따라 달라집니다.

Windows x64 ZIP 전체를 풀고 `Tablyfern.exe`를 실행하세요. 다운로드한 ZIP은 `SHA256SUMS.txt`와 대조할 수 있습니다. [제3자 고지 및 해당 소스](https://github.com/hobby-programmer-5000/tablyfern/blob/v0.2.0/THIRD-PARTY-NOTICES.ko.md)
