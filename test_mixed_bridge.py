import subprocess
import json
import re

bridge = r"D:\Bangla_Document_Converter\bijoy_bridge.mjs"

process = subprocess.Popen(
    ["node", bridge],
    stdin=subprocess.PIPE,
    stdout=subprocess.PIPE,
    stderr=subprocess.PIPE,
    text=True,
    encoding="utf-8",
    errors="strict",
    bufsize=1,
)

PROTECTED_TOKENS = ["TOR-1", "TOR-2"]


def convert_segments(text):
    pattern = "(" + "|".join(re.escape(x) for x in PROTECTED_TOKENS) + ")"

    parts = re.split(pattern, text)

    result = []

    for part in parts:
        if not part:
            continue

        if part in PROTECTED_TOKENS:
            # English token: do NOT send to Bijoy converter
            result.append(part)
        else:
            request = {
                "id": 1,
                "text": part,
            }

            process.stdin.write(
                json.dumps(request, ensure_ascii=False) + "\n"
            )
            process.stdin.flush()

            response = process.stdout.readline()
            data = json.loads(response)

            if "error" in data:
                raise RuntimeError(data["error"])

            result.append(data["text"])

    return "".join(result)


tests = [
    "mshyw³: 1) `icÎ Db¥y³KiY KwgwUi cÖwZ‡e`bmg~n (TOR-1, TOR-2)",
    "Kcvwe‡K, weD‡ev, KvßvB, iv½vgvwU",
]

for text in tests:
    print()
    print("INPUT :")
    print(text)

    output = convert_segments(text)

    print("OUTPUT:")
    print(output)

process.stdin.close()
process.wait()