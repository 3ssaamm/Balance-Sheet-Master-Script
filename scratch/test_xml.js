function colLetterToIdx(letters) {
    let col = 0;
    for (let i = 0; i < letters.length; i++) {
        col = col * 26 + (letters.charCodeAt(i) - 64);
    }
    return col - 1; // 0-based
}

function unescapeXml(str) {
    if (!str) return "";
    return str
        .replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'")
        .replace(/&#(\d+);/g, (m, dec) => String.fromCharCode(dec))
        .replace(/&#x([0-9a-fA-F]+);/g, (m, hex) => String.fromCharCode(parseInt(hex, 16)));
}

console.log("Col A:", colLetterToIdx("A")); // 0
console.log("Col B:", colLetterToIdx("B")); // 1
console.log("Col J:", colLetterToIdx("J")); // 9
console.log("Col Z:", colLetterToIdx("Z")); // 25
console.log("Col AA:", colLetterToIdx("AA")); // 26
console.log("XML unescape test:", unescapeXml("Hassan &amp; Angel &lt;test&gt; &#36;100"));
