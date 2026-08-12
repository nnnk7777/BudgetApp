const fs = require('node:fs');
const http = require('node:http');
const path = require('node:path');

const rootDirectory = path.resolve(__dirname, '..');
const files = {
    '/': 'tools/budget-chart-preview.template',
    '/budget-chart-preview.html': 'tools/budget-chart-preview.template',
    '/scripts/formatting/dailyBudgetChart.js': 'scripts/formatting/dailyBudgetChart.js'
};

http.createServer((request, response) => {
    const relativePath = files[request.url];

    if (!relativePath) {
        response.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
        response.end('Not found');
        return;
    }

    const contentType = relativePath.endsWith('.template') ? 'text/html; charset=utf-8' : 'application/javascript; charset=utf-8';
    response.writeHead(200, { 'Content-Type': contentType, 'Cache-Control': 'no-store' });
    response.end(fs.readFileSync(path.join(rootDirectory, relativePath)));
}).listen(4173, '127.0.0.1', () => {
    console.log('予算グラフプレビュー: http://127.0.0.1:4173');
});
