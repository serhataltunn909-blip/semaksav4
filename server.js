const http = require("http");
const fs = require("fs");
const XLSX = require("xlsx");

const PORT = 5000;
const EXCEL_FILE = "template.xlsx";

http.createServer((req, res) => {
  if (req.method === "POST") {
    let body = "";
    req.on("data", chunk => body += chunk);
    req.on("end", () => {
      const data = JSON.parse(body).PARAMETERS;

      const wb = XLSX.readFile(EXCEL_FILE);
      const ws = wb.Sheets[wb.SheetNames[0]];

      Object.keys(data).forEach((key, i) => {
        const row = i + 1;
        ws[`B${row}`] = { t: "n", v: data[key] };
      });

      XLSX.writeFile(wb, EXCEL_FILE);

      res.writeHead(200);
      res.end("OK");
    });
  }
}).listen(PORT);

console.log("Server çalışıyor:", PORT);
