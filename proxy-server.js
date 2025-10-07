const fs = require('fs');
const path = require('path');
const https = require('https');
const express = require('express');
const cors = require('cors');
const multer = require('multer');
const axios = require('axios');
const FormData = require('form-data');

const app = express();
const PORT = 3005;

const TARGET_URL = 'https://api-test.munichre.com/gsiuwbdim/dev/placements/v1/api/placements';

// HTTPS cert config
const httpsOptions = {
  key: fs.readFileSync(path.join(__dirname, 'certs', 'localhost-key.pem')),
  cert: fs.readFileSync(path.join(__dirname, 'certs', 'localhost-cert.pem')),
};

// Middleware
app.use(cors({
  origin: '*',
  credentials: true,
  allowedHeaders: ['Content-Type', 'Authorization', 'ocp-apim-subscription-key']
}));

const upload = multer({ storage: multer.memoryStorage() });

// Proxy handler
app.post('/', upload.single('files'), async (req, res) => {
  console.log(`➡️ Incoming POST request`);
  console.log('Headers:', req.headers);
  console.log('Form Fields:', req.body);
  console.log('File:', req.file?.originalname);

  try {
    const form = new FormData();

    // Append fields
    form.append('productCode', req.body.productCode);
    form.append('emailSender', req.body.emailSender);
    form.append('emailSubject', req.body.emailSubject);
    form.append('emailReceivedDateTime', req.body.emailReceivedDateTime);

    // Append file if exists
    if (req.file) {
      form.append('files', req.file.buffer, {
        filename: req.file.originalname,
        contentType: req.file.mimetype
      });
    }

    // Prepare headers
    const headers = {
      ...form.getHeaders(),
      'Authorization': req.headers['authorization'] || '',
      'ocp-apim-subscription-key': req.headers['ocp-apim-subscription-key'] || ''
    };

    const response = await axios.post(TARGET_URL, form, {
      headers: headers,
      httpsAgent: new https.Agent({ rejectUnauthorized: false }) // bypass cert error (local dev only)
    });

    console.log(`Backend response: ${response.status}`);
    res.status(response.status).json(response.data);

  } catch (error) {
    console.error('❌ Error forwarding request:', error.message);

    if (error.response) {
      res.status(error.response.status).json(error.response.data);
    } else {
      res.status(500).json({ error: 'Proxy Error', message: error.message });
    }
  }
});

// Start HTTPS server
https.createServer(httpsOptions, app).listen(PORT, () => {
  console.log(`Proxy server running at https://localhost:${PORT}`);
});