import readline from "node:readline";
import { ConvertToUnicode } from "bijoy-unicode-converter";

const rl = readline.createInterface({
    input: process.stdin,
    crlfDelay: Infinity,
});

rl.on("line", (line) => {
    try {
        const request = JSON.parse(line);

        const output = ConvertToUnicode(
            "bijoy",
            request.text ?? ""
        );

        // JSON is written as UTF-8 bytes directly.
        const response = JSON.stringify({
            id: request.id,
            text: output,
        });

        process.stdout.write(response + "\n");
    } catch (error) {
        const response = JSON.stringify({
            id: null,
            error: String(error),
        });

        process.stdout.write(response + "\n");
    }
});