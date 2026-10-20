function colLetterToIdx(letters) {
    let col = 0;
    for (let i = 0; i < letters.length; i++) {
        col = col * 26 + (letters.charCodeAt(i) - 64);
    }
    return col - 1;
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

function parsePeriodLine(text) {
    if (!text) return null;
    const clean = String(text).replace(/[\u2010-\u2015\u2212]/g, '-').trim();
    const match = clean.match(/period:\s*([0-9\/\.\-]+)\s*[\-–—to\s]+\s*([0-9\/\.\-]+)/i);
    if (match) {
        return { startStr: match[1], endStr: match[2] };
    }
    return null;
}

console.log("Period test:", parsePeriodLine("Period: 10/1/2026 – 10/1/2026"));
console.log("Period test with spaces:", parsePeriodLine("Period: 10/01/2026 to 10/01/2026"));
