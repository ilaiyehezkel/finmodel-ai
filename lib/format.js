export function formatMessage(text) {
  text = text.replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>");
  text = text.replace(/\*(.*?)\*/g, "<em>$1</em>");
  text = text.replace(/^### (.*$)/gm, "<h3>$1</h3>");
  text = text.replace(/^## (.*$)/gm, "<h2>$1</h2>");
  text = text.replace(/^# (.*$)/gm, "<h1>$1</h1>");
  text = text.replace(/```[\w]*\n?([\s\S]*?)```/g, "<pre><code>$1</code></pre>");
  text = text.replace(/`([^`]+)`/g, "<code>$1</code>");
  text = formatTables(text);
  text = text.replace(/^\- (.*$)/gm, "<li>$1</li>");
  text = text.replace(/(<li>.*<\/li>\n?)+/g, "<ul>$&</ul>");
  text = text.replace(/^\d+\. (.*$)/gm, "<li>$1</li>");
  text = text.replace(/^---$/gm, "<hr>");
  text = text
    .split(/\n\n+/)
    .map((p) => {
      if (
        p.startsWith("<h") ||
        p.startsWith("<ul") ||
        p.startsWith("<pre") ||
        p.startsWith("<table") ||
        p.startsWith("<hr") ||
        p.startsWith("<div")
      )
        return p;
      return `<p>${p.replace(/\n/g, "<br>")}</p>`;
    })
    .join("");
  return text;
}

function formatTables(text) {
  const tableRegex = /(\|.+\|\n)+/g;
  return text.replace(tableRegex, (match) => {
    const rows = match.trim().split("\n");
    const isHeader = (row) => /^\|[-| :]+\|$/.test(row.trim());
    let html = '<div class="table-wrap"><table>';
    let inBody = false;
    rows.forEach((row, i) => {
      if (isHeader(row)) {
        inBody = true;
        html += "<tbody>";
        return;
      }
      const parts = row.split("|");
      const cells = parts
        .filter((_, idx) => idx > 0 && idx < parts.length - 1)
        .map((c) => c.trim());
      if (i === 0) {
        html +=
          "<thead><tr>" + cells.map((c) => `<th>${c}</th>`).join("") + "</tr></thead>";
      } else {
        html += "<tr>" + cells.map((c) => `<td>${c}</td>`).join("") + "</tr>";
      }
    });
    if (inBody) html += "</tbody>";
    html += "</table></div>";
    return html;
  });
}
