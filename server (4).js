const express = require('express');
const app = express();

// Accept raw body for ALL content types
app.use((req, res, next) => {
  let data = '';
  req.setEncoding('utf8');
  req.on('data', chunk => { data += chunk; });
  req.on('end', () => {
    req.rawBody = data;
    // Try to parse as JSON
    try { req.body = JSON.parse(data); } catch(e) {
      // Try URL-decoded
      try {
        const decoded = decodeURIComponent(data);
        req.body = JSON.parse(decoded);
      } catch(e2) {
        // Try as key of URL-encoded object
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
  res.json({ status: 'ok', service: 'docx-generator', version: '5.0.0' });
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

    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'Content-Disposition': `attachment; filename="${runId}_p0_segmentation.docx"`,
      'Content-Length': docxBuffer.length
    });

    res.send(docxBuffer);

  } catch (err) {
    console.error('Error generating DOCX:', err);
    res.status(500).json({ error: err.message, stack: err.stack });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`DOCX generator service running on port ${PORT}`);
});
