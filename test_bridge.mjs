import { ConvertToUnicode } from "bijoy-unicode-converter";

const tests = [
    "cÖwZ‡e`bmg~n",
    "g~j¨vqb",
    "KwgwUi",
    "ZvwiL",
    "wbe©vnx",
];

for (const input of tests) {
    const output = ConvertToUnicode("bijoy", input);

    console.log(`INPUT : ${input}`);
    console.log(`OUTPUT: ${output}`);
    console.log("---");
}