const express = require('express');
const app = express();

app.use((req, res, next) => {
  let data = '';
  req.setEncoding('utf8');
  req.on('data', chunk => { data += chunk; });
  req.on('end', () => {
    req.rawBody = data;
    try { req.body = JSON.parse(data); } catch(e) {
      try {
        const decoded = decodeURIComponent(data);
        req.body = JSON.parse(decoded);
      } catch(e2) {
        const key = data.split('=')[0];
        try { req.body = JSON.parse(decodeURIComponent(key)); } catch(e3) {
          req.body = {};
        }
      }
    }
    next();
  });
});

const { generateP0Docx } = require('./generator');

app.get('/', (req, res) => {
  res.json({ status: 'ok', service: 'docx-generator', version: '6.0.0' });
});

app.post('/generate/p0', async (req, res) => {
  try {
    const body = req.body || {};
    const { p0Output, categoryName, runId, runLabel, totalReviews, totalProducts } = body;

    if (!p0Output) {
      return res.status(400).json({
        error: 'Missing p0Output in request body',
        receivedType: typeof body,
        receivedKeys: Object.keys(body),
        rawBodyPreview: (req.rawBody || '').slice(0, 300)
      });
    }

    const docxBuffer = await generateP0Docx(
      p0Output, categoryName, runId, runLabel, totalReviews, totalProducts
    );

    res.json({
      success: true,
      fileName: `${runId}_p0_segmentation.docx`,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      data: docxBuffer.toString('base64'),
      size: docxBuffer.length
    });

  } catch (err) {
    console.error('Error generating DOCX:', err);
    res.status(500).json({ error: err.message, stack: err.stack });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`DOCX generator service running on port ${PORT}`);
});
