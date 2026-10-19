const http = require('node:http');
const { createRuntime } = require('./gas-local.cjs');
const runtime = createRuntime();
const port = Number(process.env.PORT || 3000);
http.createServer((request, response) => {
  if (request.url !== '/') {
    response.writeHead(404).end('Not found');
    return;
  }
  try {
    const output = runtime.context.doGet();
    response.writeHead(runtime.errors.length ? 500 : 200, { 'Content-Type': 'text/html; charset=utf-8' });
    response.end(output.html);
  } catch (error) {
    console.error(error);
    response.writeHead(500).end('Local rendering failed');
  }
}).listen(port, '127.0.0.1', () => {
  console.log('Local sample renderer running on port ' + port + '. Uses initial data, not Google Sheets.');
});
