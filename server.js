const http = require('http');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { formatDocx } = require('./formatDocx');
const { verifyDocx } = require('./verify');

const port = Number(process.env.PORT) || 3000;

const publicDir = path.join(__dirname, 'public');
const uploadDir = path.join(__dirname, 'uploads');
const outputDir = path.join(__dirname, 'generated');
fs.mkdirSync(uploadDir, { recursive: true });
fs.mkdirSync(outputDir, { recursive: true });

const jobs = new Map();

function sendJson(res, code, payload) {
  res.writeHead(code, { 'Content-Type': 'application/json; charset=utf-8' });
  res.end(JSON.stringify(payload));
}

function sendFile(res, fullPath, contentType = 'text/html; charset=utf-8') {
  if (!fs.existsSync(fullPath)) {
    res.writeHead(404);
    res.end('Not Found');
    return;
  }

  res.writeHead(200, { 'Content-Type': contentType });
  fs.createReadStream(fullPath).pipe(res);
}

function parseBody(req) {
  return new Promise((resolve, reject) => {
    let body = '';
    req.on('data', (chunk) => {
      body += chunk;
      if (body.length > 50 * 1024 * 1024) {
        reject(new Error('Запит занадто великий.'));
        req.destroy();
      }
    });
    req.on('end', () => resolve(body));
    req.on('error', reject);
  });
}

function createJob(filename) {
  const id = crypto.randomUUID();
  const now = new Date().toISOString();
  const job = {
    id,
    filename,
    status: 'queued',
    progress: 0,
    log: ['Завдання створено.'],
    createdAt: now,
    updatedAt: now,
    formatStats: null,
    verifyReport: null,
    outputFilename: null,
    error: null,
  };
  jobs.set(id, job);
  return job;
}

function updateJob(job, patch) {
  Object.assign(job, patch, { updatedAt: new Date().toISOString() });
}

function pushLog(job, message, progress = null) {
  if (progress !== null) {
    job.progress = progress;
  }
  job.log.push(message);
  job.updatedAt = new Date().toISOString();
}

async function processJob(job, inputPath, outputPath, outputFilename) {
  try {
    updateJob(job, { status: 'running', progress: 5 });
    pushLog(job, 'Починаю форматування...', 10);

    const formatResult = await formatDocx(inputPath, outputPath, {
      onProgress: (message) => {
        const nextProgress = Math.min(70, job.progress + 12);
        pushLog(job, message, nextProgress);
      },
    });

    updateJob(job, { formatStats: formatResult.stats, outputFilename });
    pushLog(job, 'Форматування завершено. Запускаю перевірку...', 80);

    const verifyReport = await verifyDocx(outputPath);
    updateJob(job, { verifyReport });

    pushLog(job, verifyReport.ok ? 'Перевірка пройдена успішно.' : 'Перевірка завершена з невідповідностями.', 100);
    updateJob(job, { status: verifyReport.ok ? 'completed' : 'completed_with_warnings' });
  } catch (error) {
    updateJob(job, { status: 'failed', error: error.message, progress: 100 });
    pushLog(job, `Помилка: ${error.message}`);
  } finally {
    fs.unlink(inputPath, () => {});
  }
}

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, `http://${req.headers.host}`);

  if (req.method === 'GET' && url.pathname === '/') {
    sendFile(res, path.join(publicDir, 'index.html'));
    return;
  }

  if (req.method === 'GET' && url.pathname === '/api/jobs') {
    sendJson(res, 200, Array.from(jobs.values()));
    return;
  }

  if (req.method === 'GET' && url.pathname.startsWith('/api/jobs/')) {
    const parts = url.pathname.split('/').filter(Boolean);
    const jobId = parts[2];
    const action = parts[3];
    const job = jobs.get(jobId);

    if (!job) {
      sendJson(res, 404, { error: 'Завдання не знайдено.' });
      return;
    }

    if (action === 'download') {
      if (!job.outputFilename) {
        sendJson(res, 404, { error: 'Файл результату не знайдено.' });
        return;
      }
      const fullPath = path.join(outputDir, job.outputFilename);
      if (!fs.existsSync(fullPath)) {
        sendJson(res, 404, { error: 'Файл результату відсутній на диску.' });
        return;
      }
      res.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'Content-Disposition': `attachment; filename*=UTF-8''${encodeURIComponent(job.outputFilename)}`,
      });
      fs.createReadStream(fullPath).pipe(res);
      return;
    }

    sendJson(res, 200, job);
    return;
  }

  if (req.method === 'POST' && url.pathname === '/api/process') {
    try {
      const raw = await parseBody(req);
      const payload = JSON.parse(raw || '{}');

      if (!payload.filename || !payload.contentBase64) {
        sendJson(res, 400, { error: 'Передайте filename та contentBase64.' });
        return;
      }
      if (!payload.filename.toLowerCase().endsWith('.docx')) {
        sendJson(res, 400, { error: 'Дозволені лише файли .docx' });
        return;
      }

      const safeName = payload.filename.replace(/[^\p{L}\p{N}._-]+/gu, '_');
      const inputFilename = `${Date.now()}-${safeName}`;
      const inputPath = path.join(uploadDir, inputFilename);
      fs.writeFileSync(inputPath, Buffer.from(payload.contentBase64, 'base64'));

      const job = createJob(payload.filename);
      const outputFilename = `${path.parse(payload.filename).name}-formatted-${job.id}.docx`;
      const outputPath = path.join(outputDir, outputFilename);

      sendJson(res, 202, { jobId: job.id });
      processJob(job, inputPath, outputPath, outputFilename);
      return;
    } catch (error) {
      sendJson(res, 400, { error: error.message || 'Некоректний запит.' });
      return;
    }
  }

  res.writeHead(404);
  res.end('Not Found');
});

server.listen(port, () => {
  console.log(`UI доступний на http://localhost:${port}`);
});
