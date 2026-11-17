# Tablyfern 0.3.0

The Windows workspace now uses neutral surfaces, a blue save action and clearer type. Input and preview sit side by side, with a vertical layout for smaller windows. Opening files previews them automatically. Pasting or editing input reveals **Update preview**, also available with **Ctrl+Enter**. Both English and Korean README demos show the new workspace.

This release also includes the reviewed fixes for numeric precision decisions, concurrent iteration over disk-backed tables, UTF-8 sizing of large pasted input, ODS row padding, file-lock reporting and temporary-file cleanup. Running reads and exports have a **Cancel** action. Closing during a job offers cancellation before quitting.

Extract the entire Windows x64 ZIP and run `Tablyfern.exe`, keeping `_internal` and `assets` alongside it. The ZIP includes synthetic samples, bilingual documentation, the Lucide and Feather icon notices and dependency licenses. Compare its SHA-256 with `SHA256SUMS.txt`. [Third-party notices](https://github.com/hobby-programmer-5000/tablyfern/blob/v0.3.0/THIRD-PARTY-NOTICES.md)

Large inputs still depend on available memory and temporary/output disk space. Each input file maps to one worksheet, subject to worksheet and output-package limits. The **Limits** dialog explains capacity and format details.

## 한국어

Windows 화면에 중립색 배경과 파란 저장 버튼을 적용하고 글자를 더 읽기 쉽게 다듬었습니다. 입력과 표를 나란히 표시하며 작은 창에서는 위아래로 배치합니다. 파일을 열면 자동으로 미리 보기가 만들어지고, 데이터를 붙여넣거나 수정하면 **미리보기 갱신** 버튼이 나타납니다. **Ctrl+Enter**도 사용할 수 있습니다. 한글·영문 README 데모를 새 화면으로 갱신했습니다.

숫자 정밀도 판단, 디스크 기반 표의 동시 읽기, 큰 붙여넣기의 UTF-8 크기 판별, ODS 행 너비, 파일 잠금 오류 안내, 임시 파일 정리에 대한 검토된 수정도 포함했습니다. 읽기와 저장 중 **취소**를 사용할 수 있고, 작업 중 창을 닫으면 취소 후 종료할지 확인합니다.

Windows x64 ZIP 전체를 풀고 `Tablyfern.exe`를 실행하세요. `_internal`과 `assets` 폴더를 EXE와 함께 두세요. 가상 샘플, 한글·영문 문서, Lucide·Feather 아이콘 고지와 의존성 라이선스가 포함됩니다. ZIP의 SHA-256은 `SHA256SUMS.txt`와 대조할 수 있습니다. [제3자 고지](https://github.com/hobby-programmer-5000/tablyfern/blob/v0.3.0/THIRD-PARTY-NOTICES.ko.md)

대용량 입력은 가용 메모리와 임시·출력 디스크 공간의 영향을 받습니다. 입력 파일마다 시트 한 장을 만들며 시트 및 출력 패키지 한도가 적용됩니다. 앱의 **사용 한도**에서 용량과 형식별 설명을 확인할 수 있습니다.
