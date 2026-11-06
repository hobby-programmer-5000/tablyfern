# Sample data

The [English samples](en/) are a self-contained set. Select all three files together to create an `orders` tab with English city names, an `events` tab from JSON Lines, and a `precision` tab for long IDs and formula-like text.

Select `orders.json`, `events.jsonl` and `precision.json` together to try a three-tab workbook. Add `empty.json` to check that a header-only tab is retained.

These are original synthetic examples with invented order IDs and event records. They are included under the repository's [MIT license](../LICENSE). Run `python fixtures.py .cache/fixtures` for the larger fixture catalog, which also contains deliberately invalid inputs.

# 샘플 데이터

[영문 샘플](en/)의 파일 3개를 함께 선택하면 영문 도시명이 들어 있는 `orders` 탭, JSONL에서 만든 `events` 탭, 긴 ID와 수식처럼 보이는 문자열을 확인할 수 있는 `precision` 탭을 만들 수 있습니다.

`orders.json`, `events.jsonl`, `precision.json`을 함께 선택하면 탭 3개짜리 통합 문서를 만들 수 있습니다. `empty.json`을 추가하면 머리글만 있는 탭도 만들어지는지 확인할 수 있습니다.

주문 번호와 이벤트를 가상으로 만든 자체 예제입니다. 저장소의 [MIT 라이선스](../LICENSE)를 적용합니다. `python fixtures.py .cache/fixtures`로 의도적으로 잘못된 입력까지 포함한 전체 테스트 데이터를 생성할 수 있습니다.
