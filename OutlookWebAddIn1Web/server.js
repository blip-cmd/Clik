const https = require('https');
const fs = require('fs');
const path = require('path');

const options = {
    key: fs.readFileSync('localhost-key.pem'),
    cert: fs.readFileSync('localhost.pem')
};

https.createServer(options, (req, res) => {
    const cleanUrl = req.url.split('?')[0];
    const filePath = path.join(__dirname, cleanUrl === '/' ? 'MessageRead.html' : cleanUrl);

    console.log('🔍 Requested:', cleanUrl);
    console.log('📁 File path:', filePath);
    console.log('📂 File exists:', fs.existsSync(filePath));

    fs.readFile(filePath, (err, data) => {
        if (err) {
            console.log('❌ Error:', err.message);
            res.writeHead(404);
            res.end('File not found');
        } else {
            res.writeHead(200, { 'Content-Type': 'text/html' });
            res.end(data);
        }
    });
}).listen(44300, () => {
    console.log('✅ HTTPS server running at https://localhost:44300');
});