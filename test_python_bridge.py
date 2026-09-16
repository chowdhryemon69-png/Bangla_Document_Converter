import subprocess
import json

bridge = r"D:\Bangla_Document_Converter\bijoy_bridge.mjs"

process = subprocess.Popen(
    ["node", bridge],
    stdin=subprocess.PIPE,
    stdout=subprocess.PIPE,
    text=True,
    encoding="utf-8",
    errors="strict",
    bufsize=1,
)

tests = [
    "cÖwZ‡e`bmg~n",
    "g~j¨vqb",
    "KwgwUi",
]

for i, text in enumerate(tests, start=1):

    request = {
        "id": i,
        "text": text,
    }

    process.stdin.write(
        json.dumps(
            request,
            ensure_ascii=False,
        ) + "\n"
    )

    process.stdin.flush()

    response = process.stdout.readline()

    print(response.rstrip())

process.stdin.close()
process.wait()