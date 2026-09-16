import { ConvertToUnicode } from "bijoy-unicode-converter";

const tests = [
    "cÖwZ‡e`bmg~n",
    "g~j¨vqb",
    "KwgwUi",
    "ZvwiL",
    "wbe©vnx",
    "hvwš¿K msiÿY wefvM-1",
    "Kcvwe‡K, weD‡ev, KvßvB, iv½vgvwU",
    "mshyw³: 1) `icÎ Db¥y³KiY KwgwUi cÖwZ‡e`bmg~n (TOR-1, TOR-2)"
];

for (const input of tests) {
    const output = ConvertToUnicode("bijoy", input);

    console.log("--------------------------------------------------");
    console.log("INPUT : ", input);
    console.log("OUTPUT: ", output);
}