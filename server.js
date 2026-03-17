const express = require('express');
const app = express();
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));
app.use(express.text({ type: '*/*', limit: '50mb' }));

const { generateP0Docx } = require('./generator');

app.get('/', (req, res) => {
  res.json({ status: 'ok', service: 'docx-generator', version: '4.0.0' });
});

app.post('/generate/p0', async (req, res) => {
  try {
    let body = req.body;

    // Handle n8n sending JSON as a URL-encoded key
    if (typeof body === 'object' && body !== null) {
      const keys = Object.keys(body);
      if (keys.length === 1 && keys[0] !== 'p0Output') {
        try { body = JSON.parse(keys[0]); } catch(e) {}
      }
    }

    // Handle string body
    if (typeof body === 'string') {
      try { body = JSON.parse(body); } catch(e) {}
    }

    const { p0Output, categoryName, runId, runLabel, totalReviews, totalProducts } = body || {};

if (!p0Output) {
  return res.status(400).json({
    error: 'Missing p0Output in request body',
    receivedType: typeof body,
    receivedKeys: body ? Object.keys(body) : 'null',
    bodyPreview: JSON.stringify(body).slice(0, 500)
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
