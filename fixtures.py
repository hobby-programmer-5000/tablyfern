"""Generate deterministic synthetic JSON and JSONL files for testing."""

import argparse
import json
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class Case:
    filename: str
    source: str
    expected: tuple[int, int] | str

    @property
    def mode(self) -> str:
        return "json" if self.filename.endswith(".json") else "jsonl"


def catalog() -> list[Case]:
    events = [
        {"id": index, "type": ("view", "click", "purchase")[index % 3],
         "amount": index / 10, "actor": {"region": ("KR", "US")[index % 2]}}
        for index in range(3000)
    ]
    return [
        Case("01-records.json", '[{"name":"Mina","score":7},{"name":"Jin","score":9}]', (2, 2)),
        Case("02-single-object.json", '{"product":"notebook","stock":0,"active":false}', (1, 3)),
        Case("03-nested-object.json", '{"a":{"b":{"c":1}},"address":{"city":"서울"},"ok":true}', (1, 3)),
        Case("04-path-collisions.json", r'{"a.b":1,"a":{"b":2},"a\\b":3,"":4,"\\0":5}', (1, 5)),
        Case("05-sparse-records.json", '[{"first":1,"common":null},{"last":2},{"common":3},{}]', (4, 3)),
        Case("06-nested-arrays.json", '[{"tags":["red","blue"],"matrix":[[1],[2,3]],"blank":{}}]', (1, 3)),
        Case("07-top-level-array.json", '[["a",2],[false,null]]', (2, 1)),
        Case("08-mixed-values.json", '[{"id":1},null,true,"007",[4],-5.5]', (6, 1)),
        Case("09-scalar-string.json", '"문자열 🚀"', (1, 1)),
        Case("10-scalar-number.json", '-12.375', (1, 1)),
        Case("11-scalar-null.json", 'null', (1, 1)),
        Case("12-empty-object.json", '{}', (1, 1)),
        Case("13-empty-array.json", '[]', (0, 1)),
        Case("14-types-unicode.json", '{"ko":"서울","jp":"東京","ar":"مرحبا","emoji":"👩🏽‍💻","empty":"","none":null,"on":true}', (1, 7)),
        Case("15-numeric-precision.json", '{"id":9007199254740993,"money":0.1234567890123456,"small":1e-400,"big":1e400,"ok":1e2,"zero":0}', (1, 6)),
        Case("16-formula-and-url.json", '{"=1+1":"=2+2","at":" @SUM(1)","wide":"＝1+1","minus":"-42","url":"https://example.com","number":-42}', (1, 6)),
        Case("17-delimiters.json", r'{"comma":"a,b","quotes":"\"test\"","tab":"x\ty","lines":"one\r\ntwo","spaces":"  a  b  "}', (1, 5)),
        Case("18-bom.json", '\ufeff{"name":"BOM"}', (1, 1)),
        Case("19-wide.json", json.dumps({f"field_{i:03}": i for i in range(300)}), (1, 300)),
        Case("20-cell-boundary.json", json.dumps({"text": "a" * 32767}), (1, 1)),
        Case("21-events.jsonl", "\n".join(map(json.dumps, events)) + "\n", (3000, 4)),
        Case("22-object-lines.jsonl", '{"a":1}\n{"b":2}\n{"a":3,"b":4}\n', (3, 2)),
        Case("23-crlf-bom.ndjson", '\ufeff{"city":"부산"}\r\n{"city":"서울"}\r\n', (2, 1)),
        Case("24-mixed-lines.jsonl", '{"a":1}\n"한글"\nfalse\nnull\n[1,2]', (5, 1)),
        Case("25-single-array-line.jsonl", '[1,2,3]\n', (1, 1)),
        Case("26-truncated.json", '{"a":', "invalid"),
        Case("27-duplicate.json", '{"x":1,"x":2}', "duplicate"),
        Case("28-non-json-number.json", '[NaN]', "constant"),
        Case("29-invalid-unicode.json", r'"\ud800"', "unicode"),
        Case("30-overlong-cell.json", json.dumps("a" * 32768), "cell_limit"),
        Case("31-empty-line.jsonl", '{"a":1}\n\n{"a":2}\n', "invalid"),
        Case("32-malformed-line.jsonl", '{"a":1}\n{"a":}\n', "invalid"),
        Case("33-duplicate-line.jsonl", '{"a":1}\n{"b":1,"b":2}\n', "duplicate_line"),
        Case("34-infinite-line.jsonl", '{"a":1}\nInfinity\n', "constant_line"),
        Case("35-pretty-document.jsonl", '[\n  1,\n  2\n]\n', "invalid"),
    ]


def generate(output: Path, cases: list[Case]) -> None:
    output.mkdir(parents=True)
    for case in cases:
        (output / case.filename).write_bytes(case.source.encode("utf-8"))


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output", type=Path, help="a new directory for the synthetic inputs")
    args = parser.parse_args()
    generate(args.output, catalog())
