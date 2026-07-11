function parsePipeResponse(text) {
    return cleanAiResponseText(text)
        .split("\n")
        .map(function (line) {
            return line.trim();
        })
        .filter(function (line) {
            return /^\d+\|/.test(line);
        })
        .map(function (line) {
            var separatorIndex = line.indexOf("|");
            return {
                index: parseInt(line.substring(0, separatorIndex), 10),
                cleanedMemo: line.substring(separatorIndex + 1).trim()
            };
        })
        .filter(function (item) {
            return !isNaN(item.index);
        });
}

function cleanAiResponseText(text) {
    return String(text)
        .replace(/```[\s\S]*?\n/g, "")
        .replace(/```/g, "")
        .trim();
}
