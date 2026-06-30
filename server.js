require('dotenv').config();

const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const multer = require('multer');
const XLSX = require('xlsx-js-style');
const nodemailer = require('nodemailer');
const cron = require('node-cron');

const app = express();

const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

const dataDir = path.join(__dirname, 'data');
const uploadsDir = path.join(__dirname, 'uploads');

if (!fs.existsSync(dataDir)) fs.mkdirSync(dataDir);
if (!fs.existsSync(uploadsDir)) fs.mkdirSync(uploadsDir);

app.use('/uploads', express.static(uploadsDir));

const files = {
  entries: path.join(dataDir, 'entries.json'),
  products: path.join(dataDir, 'products.json'),
  numbers: path.join(dataDir, 'numbers.json'),
  merchants: path.join(dataDir, 'merchants.json'),
  profile: path.join(dataDir, 'profile.json'),
};

function readJson(file, fallback = []) {
  try {
    if (!fs.existsSync(file)) {
      fs.writeFileSync(file, JSON.stringify(fallback, null, 2));
      return fallback;
    }

    const content = fs.readFileSync(file, 'utf8');
    if (!content.trim()) return fallback;

    return JSON.parse(content);
  } catch {
    return fallback;
  }
}

function writeJson(file, data) {
  fs.writeFileSync(file, JSON.stringify(data, null, 2));
}

function getIndiaTime() {
  return new Date().toLocaleTimeString('en-IN', {
    timeZone: 'Asia/Kolkata',
    hour: '2-digit',
    minute: '2-digit',
    hour12: true,
  });
}

function normalizeNumber(value) {
  return String(value || '')
    .replace('whatsapp:', '')
    .replace('+', '')
    .replace(/[^0-9]/g, '')
    .trim();
}

function last10(value) {
  return normalizeNumber(value).slice(-10);
}

function findMerchantBySender(sender, merchants) {
  const senderFull = normalizeNumber(sender);
  const senderLast = last10(sender);

  return merchants.find((merchant) => {
    const merchantFull = normalizeNumber(merchant.merchant_number);
    const merchantLast = last10(merchant.merchant_number);

    return (
      merchantFull === senderFull ||
      merchantFull === senderLast ||
      merchantLast === senderLast
    );
  });
}

function getIndiaMinutes() {
  const now = new Date();
  const indiaTime = new Date(
    now.toLocaleString('en-US', { timeZone: 'Asia/Kolkata' })
  );

  return indiaTime.getHours() * 60 + indiaTime.getMinutes();
}

function timeToMinutes(value) {
  const [hours, minutes] = String(value || '00:00').split(':').map(Number);
  if (Number.isNaN(hours) || Number.isNaN(minutes)) return 0;
  return hours * 60 + minutes;
}

function isWithinTime(startTime, endTime) {
  const now = getIndiaMinutes();
  const start = timeToMinutes(startTime || '00:00');
  const end = timeToMinutes(endTime || '23:59');

  if (start <= end) return now >= start && now <= end;

  return now >= start || now <= end;
}

function twilioReply(res, message) {
  res.set('Content-Type', 'text/xml');

  return res.status(200).send(`
<Response>
  <Message>${message}</Message>
</Response>
`);
}

function parseWhatsappMessage(message) {
  const name = message.match(/Name:\s*(.*)/i)?.[1]?.trim() || '';
  const product = message.match(/Product:\s*(.*)/i)?.[1]?.trim() || '';
  const quantity = message.match(/Quantity:\s*(.*)/i)?.[1]?.trim() || '';
  const time = message.match(/Time:\s*(.*)/i)?.[1]?.trim() || '';

  return { name, product, quantity, time };
}

function extractQuantityNumber(quantity) {
  const match = String(quantity || '').match(/[\d.]+/);
  return match ? Number(match[0]) : 0;
}

function buildStats(entries) {
  const today = new Date().toISOString().slice(0, 10);

  const todayEntries = entries.filter((entry) => {
    const created = entry.created_at || entry.createdAt || '';
    return String(created).slice(0, 10) === today;
  });

  const productSummary = {};

  entries.forEach((entry) => {
    const productName = entry.product || 'Unknown';
    productSummary[productName] =
      (productSummary[productName] || 0) + extractQuantityNumber(entry.quantity);
  });

  return {
    trackedToday: todayEntries.length,
    totalLoads: entries.length,
    totalQuantity: entries.reduce(
      (sum, entry) => sum + extractQuantityNumber(entry.quantity),
      0
    ),
    productSummary,
  };
}

function getNextProductNumber(products) {
  const used = new Set();

  products.forEach((product) => {
    const num = Number(product.product_number);
    if (!Number.isNaN(num) && num >= 0 && num <= 99) used.add(num);
  });

  for (let i = 0; i <= 99; i += 1) {
    if (!used.has(i)) return i;
  }

  return null;
}

function resolveProductNumber(products, inputNumber) {
  if (inputNumber !== undefined && inputNumber !== null && inputNumber !== '') {
    const num = Number(inputNumber);

    if (Number.isNaN(num) || num < 0 || num > 99) {
      return { error: 'Product number must be between 0 and 99' };
    }

    const exists = products.some((p) => Number(p.product_number) === num);

    if (exists) {
      return { error: `Product number ${num} already exists` };
    }

    return { value: num };
  }

  const next = getNextProductNumber(products);

  if (next === null) {
    return { error: 'All product numbers 0-99 are already used' };
  }

  return { value: next };
}

function formatDateDDMMYYYY(value) {
  if (!value) return '';

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return '';

  const dd = String(date.getDate()).padStart(2, '0');
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const yyyy = date.getFullYear();

  return `${dd}-${mm}-${yyyy}`;
}

function createExcelBuffer(entries) {
  const rows = entries.map((entry, index) => ({
    No: index + 1,
    Sender: entry.merchant_name || entry.name || '',
    code: entry.product_number ?? entry.product ?? '',
    Quantity: extractQuantityNumber(entry.quantity),
    Time: entry.time || entry.received_time || '',
    'Date(dd-mm-yyyy)': formatDateDDMMYYYY(entry.created_at || entry.createdAt),
  }));

  while (rows.length < 15) {
    rows.push({
      No: '',
      Sender: '',
      code: '',
      Quantity: '',
      Time: '',
      'Date(dd-mm-yyyy)': '',
    });
  }

  const sheet = XLSX.utils.json_to_sheet(rows, {
    header: ['No', 'Sender', 'code', 'Quantity', 'Time', 'Date(dd-mm-yyyy)'],
  });

  sheet['!cols'] = [
    { wch: 6 },
    { wch: 18 },
    { wch: 10 },
    { wch: 14 },
    { wch: 14 },
    { wch: 20 },
  ];

  const range = XLSX.utils.decode_range(sheet['!ref']);

  for (let row = range.s.r; row <= range.e.r; row++) {
    for (let col = range.s.c; col <= range.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });

      if (!sheet[cellAddress]) {
        sheet[cellAddress] = { t: 's', v: '' };
      }

      sheet[cellAddress].s = {
        border: {
          top: { style: 'thin', color: { rgb: '000000' } },
          bottom: { style: 'thin', color: { rgb: '000000' } },
          left: { style: 'thin', color: { rgb: '000000' } },
          right: { style: 'thin', color: { rgb: '000000' } },
        },
        alignment: {
          horizontal: row === 0 ? 'center' : 'left',
          vertical: 'center',
        },
        font: {
          bold: row === 0,
        },
      };
    }
  }

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, 'Merchant Data');

  return XLSX.write(workbook, {
    type: 'buffer',
    bookType: 'xlsx',
  });
}

function getPreviousMonthEntries(entries) {
  const now = new Date();
  const firstDayThisMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  const firstDayPreviousMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);

  return entries.filter((entry) => {
    const date = new Date(entry.created_at || entry.createdAt || '');
    return date >= firstDayPreviousMonth && date < firstDayThisMonth;
  });
}

function getMailTransporter() {
  return nodemailer.createTransport({
    host: process.env.SMTP_HOST || 'smtp.gmail.com',
    port: Number(process.env.SMTP_PORT || 465),
    secure: true,
    auth: {
      user: process.env.SMTP_USER,
      pass: process.env.SMTP_PASS,
    },
  });
}

/* ---------------- BASIC ---------------- */

app.get('/', (req, res) => {
  res.send('Backend running');
});

/* ---------------- IMAGE UPLOAD ---------------- */

const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, uploadsDir);
  },
  filename: (req, file, cb) => {
    const safeName = file.originalname.replace(/\s+/g, '-');
    cb(null, `${Date.now()}-${safeName}`);
  },
});

const upload = multer({ storage });

app.post('/upload/image', upload.single('image'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'Image is required' });
  }

  const imageUrl = `https://merchant-backend-production-591b.up.railway.app/uploads/${req.file.filename}`;

  return res.json({ imageUrl });
});

/* ---------------- PRODUCTS ---------------- */

app.get('/products', (req, res) => {
  res.json(readJson(files.products, []));
});

app.post('/products', (req, res) => {
  const products = readJson(files.products, []);
  const { name, title, image, product_number } = req.body;

  if (!name || !title) {
    return res.status(400).json({
      error: 'Product name and title are required',
    });
  }

  const resolved = resolveProductNumber(products, product_number);

  if (resolved.error) {
    return res.status(400).json({ error: resolved.error });
  }

  const product = {
    id: String(Date.now()),
    name,
    title,
    image: image || '',
    product_number: resolved.value,
    created_at: new Date().toISOString(),
  };

  products.unshift(product);
  writeJson(files.products, products);

  return res.json(product);
});

app.put('/products/:id', (req, res) => {
  let products = readJson(files.products, []);
  const id = String(req.params.id);
  const existing = products.find((item) => String(item.id) === id);

  if (!existing) {
    return res.status(404).json({ error: 'Product not found' });
  }

  let nextProductNumber = existing.product_number;

  if (
    req.body.product_number !== undefined &&
    req.body.product_number !== null &&
    req.body.product_number !== ''
  ) {
    const num = Number(req.body.product_number);

    if (Number.isNaN(num) || num < 0 || num > 99) {
      return res.status(400).json({ error: 'Product number must be between 0 and 99' });
    }

    const duplicate = products.some(
      (item) => String(item.id) !== id && Number(item.product_number) === num
    );

    if (duplicate) {
      return res.status(400).json({ error: `Product number ${num} already exists` });
    }

    nextProductNumber = num;
  }

  products = products.map((item) =>
    String(item.id) === id
      ? {
          ...item,
          name: req.body.name ?? item.name,
          title: req.body.title ?? item.title,
          image: req.body.image ?? item.image,
          product_number: nextProductNumber,
        }
      : item
  );

  writeJson(files.products, products);

  return res.json({ message: 'Product updated successfully' });
});

app.delete('/products/:id', (req, res) => {
  let products = readJson(files.products, []);
  const id = String(req.params.id);

  products = products.filter((item) => String(item.id) !== id);
  writeJson(files.products, products);

  return res.json({ message: 'Product deleted successfully' });
});

app.get('/products/number/:number', (req, res) => {
  const products = readJson(files.products, []);
  const number = Number(req.params.number);

  const product = products.find((item) => Number(item.product_number) === number);

  if (!product) {
    return res.status(404).json({ error: 'Product not found' });
  }

  return res.json(product);
});

/* ---------------- ENTRIES ---------------- */

app.get('/entries', (req, res) => {
  res.json(readJson(files.entries, []));
});

app.post('/api/entries', (req, res) => {
  const entries = readJson(files.entries, []);
  const products = readJson(files.products, []);

  const productNumber =
    req.body.product_number !== undefined && req.body.product_number !== ''
      ? Number(req.body.product_number)
      : null;

  const matchedProduct = products.find(
    (p) =>
      Number(p.product_number) === productNumber ||
      String(p.name || '').toLowerCase() === String(req.body.product || '').toLowerCase()
  );

  const entry = {
    id: String(Date.now()),
    sender: req.body.sender || 'manual-entry',
    receiver: req.body.receiver || '',
    merchant_name: req.body.merchant_name || req.body.name || '',
    name: req.body.name || req.body.merchant_name || '',
    product: req.body.product || matchedProduct?.name || '',
    product_number:
      productNumber !== null ? productNumber : matchedProduct?.product_number ?? '',
    quantity: req.body.quantity || '',
    time: req.body.time || getIndiaTime(),
    raw_message: req.body.raw_message || '',
    createdAt: new Date().toISOString(),
    created_at: new Date().toISOString(),
  };

  entries.unshift(entry);
  writeJson(files.entries, entries);

  return res.json(entry);
});

app.get('/entries/product/:number', (req, res) => {
  const entries = readJson(files.entries, []);
  const number = Number(req.params.number);

  const filtered = entries.filter((entry) => Number(entry.product_number) === number);

  return res.json({
    productNumber: number,
    total: filtered.length,
    stats: buildStats(filtered),
    data: filtered,
  });
});

app.get('/entries/merchant/:name', (req, res) => {
  const entries = readJson(files.entries, []);
  const name = String(req.params.name || '').toLowerCase();

  const filtered = entries.filter((entry) =>
    String(entry.merchant_name || entry.name || '').toLowerCase().includes(name)
  );

  return res.json({
    merchant: req.params.name,
    total: filtered.length,
    stats: buildStats(filtered),
    data: filtered,
  });
});

/* ---------------- STATS ---------------- */

app.get('/stats', (req, res) => {
  const entries = readJson(files.entries, []);
  res.json(buildStats(entries));
});

/* ---------------- WHATSAPP WEBHOOK ---------------- */

app.get('/webhook', (req, res) => {
  res.send('Webhook route is available. Twilio must use POST.');
});

app.post('/webhook', (req, res) => {
  const message = req.body.Body || '';
  const sender = req.body.From || '';
  const receiver = req.body.To || '';

  const merchants = readJson(files.merchants, []);
  const merchantConfig = findMerchantBySender(sender, merchants);

  if (!merchantConfig) {
    return twilioReply(
      res,
      'Your number is not configured for tracking. Please contact admin.'
    );
  }

  if (merchantConfig.status !== 'active') {
    return twilioReply(res, 'Your account is deactivated. Please contact admin.');
  }

  const allowed = isWithinTime(
    merchantConfig.tracking_start_time || '00:00',
    merchantConfig.tracking_end_time || '23:59'
  );

  if (!allowed) {
    return twilioReply(res, 'Your time period was completed. Please contact admin.');
  }

  const parsed = parseWhatsappMessage(message);

  if (!parsed.name && !parsed.product && !parsed.quantity && !parsed.time) {
    return twilioReply(res, 'Invalid format. Please send Name, Product, Quantity.');
  }

  const products = readJson(files.products, []);
  const entries = readJson(files.entries, []);

  const productInput = parsed.product;
  const productNumberFromMessage = Number(productInput);

  let matchedProduct = null;
  let finalProductNumber = '';

  if (!Number.isNaN(productNumberFromMessage)) {
    matchedProduct = products.find(
      (p) => Number(p.product_number) === productNumberFromMessage
    );
  }

  if (!matchedProduct) {
    matchedProduct = products.find(
      (p) => String(p.name || '').toLowerCase() === String(productInput || '').toLowerCase()
    );
  }

  if (matchedProduct) {
    finalProductNumber = matchedProduct.product_number;
  } else if (!Number.isNaN(productNumberFromMessage)) {
    finalProductNumber = productNumberFromMessage;
  }

  const receivedTime = getIndiaTime();

  const entry = {
    id: String(Date.now()),
    sender,
    receiver,
    merchant_id: merchantConfig.id,
    merchant_name: merchantConfig.merchant_name || parsed.name,
    merchant_number: merchantConfig.merchant_number,
    name: parsed.name || merchantConfig.merchant_name,
    product: matchedProduct?.name || parsed.product,
    product_number: finalProductNumber,
    quantity: parsed.quantity,
    time: parsed.time || receivedTime,
    received_time: receivedTime,
    raw_message: message,
    createdAt: new Date().toISOString(),
    created_at: new Date().toISOString(),
  };

  entries.unshift(entry);
  writeJson(files.entries, entries);

  return twilioReply(res, 'Entry recorded successfully ✅');
});

/* ---------------- BUSINESS NUMBERS ---------------- */

app.get('/numbers', (req, res) => {
  res.json(readJson(files.numbers, []));
});

app.get('/business-numbers', (req, res) => {
  res.json(readJson(files.numbers, []));
});

app.post('/numbers', (req, res) => {
  const numbers = readJson(files.numbers, []);
  const { label, phone_number } = req.body;

  if (!label || !phone_number) {
    return res.status(400).json({
      error: 'Label and phone number are required',
    });
  }

  const item = {
    id: String(Date.now()),
    label,
    phone_number: normalizeNumber(phone_number),
    status: 'active',
    created_at: new Date().toISOString(),
  };

  numbers.unshift(item);
  writeJson(files.numbers, numbers);

  return res.json(item);
});

app.put('/numbers/:id/status', (req, res) => {
  let numbers = readJson(files.numbers, []);
  const id = String(req.params.id);

  numbers = numbers.map((item) =>
    String(item.id) === id ? { ...item, status: req.body.status || item.status } : item
  );

  writeJson(files.numbers, numbers);

  return res.json({ message: 'Business number updated successfully' });
});

app.delete('/numbers/:id', (req, res) => {
  let numbers = readJson(files.numbers, []);
  const id = String(req.params.id);

  numbers = numbers.filter((item) => String(item.id) !== id);
  writeJson(files.numbers, numbers);

  return res.json({ message: 'Business number deleted successfully' });
});

/* ---------------- MERCHANT TRACKING ---------------- */

app.get('/tracked-merchants', (req, res) => {
  res.json(readJson(files.merchants, []));
});

app.post('/tracked-merchants', (req, res) => {
  const merchants = readJson(files.merchants, []);
  const { merchant_name, merchant_number, tracking_start_time, tracking_end_time } = req.body;

  if (!merchant_name || !merchant_number) {
    return res.status(400).json({
      error: 'Merchant name and number are required',
    });
  }

  const merchant = {
    id: String(Date.now()),
    merchant_name,
    merchant_number: normalizeNumber(merchant_number),
    tracking_start_time: tracking_start_time || '00:00',
    tracking_end_time: tracking_end_time || '23:59',
    status: 'active',
    created_at: new Date().toISOString(),
  };

  merchants.unshift(merchant);
  writeJson(files.merchants, merchants);

  return res.json(merchant);
});

app.put('/tracked-merchants/:id/status', (req, res) => {
  let merchants = readJson(files.merchants, []);
  const id = String(req.params.id);

  merchants = merchants.map((item) =>
    String(item.id) === id ? { ...item, status: req.body.status || item.status } : item
  );

  writeJson(files.merchants, merchants);

  return res.json({ message: 'Merchant status updated successfully' });
});

app.put('/tracked-merchants/:id/schedule', (req, res) => {
  let merchants = readJson(files.merchants, []);
  const id = String(req.params.id);

  merchants = merchants.map((item) =>
    String(item.id) === id
      ? {
          ...item,
          tracking_start_time: req.body.tracking_start_time || item.tracking_start_time,
          tracking_end_time: req.body.tracking_end_time || item.tracking_end_time,
        }
      : item
  );

  writeJson(files.merchants, merchants);

  return res.json({ message: 'Merchant schedule updated successfully' });
});

app.delete('/tracked-merchants/:id', (req, res) => {
  let merchants = readJson(files.merchants, []);
  const id = String(req.params.id);

  merchants = merchants.filter((item) => String(item.id) !== id);
  writeJson(files.merchants, merchants);

  return res.json({ message: 'Merchant deleted successfully' });
});

/* ---------------- PROFILE ---------------- */

app.get('/profile', (req, res) => {
  res.json(readJson(files.profile, {}));
});

app.put('/profile', (req, res) => {
  const profile = {
    name: req.body.name || '',
    phone: req.body.phone || '',
    email: req.body.email || '',
    updated_at: new Date().toISOString(),
  };

  writeJson(files.profile, profile);

  return res.json({
    message: 'Profile updated successfully',
    profile,
  });
});

app.put('/profile/password', (req, res) => {
  return res.json({
    message: 'Password updated successfully',
  });
});

/* ---------------- PRODUCT DAILY REPORT ---------------- */

app.get('/reports/product/:number/daily', (req, res) => {
  const entries = readJson(files.entries, []);
  const number = Number(req.params.number);

  const filtered = entries.filter((entry) => Number(entry.product_number) === number);

  const dailyMap = {};

  filtered.forEach((entry) => {
    const date = String(entry.created_at || entry.createdAt || '').slice(0, 10);

    if (!dailyMap[date]) {
      dailyMap[date] = {
        date,
        entries: 0,
        totalQuantity: 0,
        data: [],
      };
    }

    dailyMap[date].entries += 1;
    dailyMap[date].totalQuantity += extractQuantityNumber(entry.quantity);
    dailyMap[date].data.push(entry);
  });

  const daily = Object.values(dailyMap).sort((a, b) => b.date.localeCompare(a.date));

  return res.json({
    productNumber: number,
    totalEntries: filtered.length,
    daily,
  });
});

/* ---------------- EXPORT ---------------- */

app.get('/export/excel/:type', (req, res) => {
  const entries = readJson(files.entries, []);
  const buffer = createExcelBuffer(entries);

  res.setHeader('Content-Disposition', 'attachment; filename=merchant-data.xlsx');
  res.setHeader(
    'Content-Type',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  );

  return res.send(buffer);
});

/* ---------------- EMAIL ---------------- */

app.post('/reports/send-monthly-test', async (req, res) => {
  try {
    const emailTo = req.body.report_email || process.env.REPORT_EMAIL_TO;

    if (!emailTo) {
      return res.status(400).json({
        error: 'Report email is required',
      });
    }

    const entries = readJson(files.entries, []);
    const buffer = createExcelBuffer(entries);

    const transporter = getMailTransporter();

    await transporter.sendMail({
      from: process.env.SMTP_USER,
      to: emailTo,
      subject: 'Merchant Monthly Report',
      text: 'Please find attached the merchant monthly report.',
      attachments: [
        {
          filename: 'merchant-monthly-report.xlsx',
          content: buffer,
        },
      ],
    });

    return res.json({
      message: 'Monthly report email sent successfully',
      totalEntries: entries.length,
    });
  } catch {
    return res.status(500).json({
      error: 'Failed to send monthly report',
    });
  }
});

cron.schedule(
  '0 9 1 * *',
  async () => {
    try {
      const emailTo = process.env.REPORT_EMAIL_TO;
      if (!emailTo) return;

      const allEntries = readJson(files.entries, []);
      const previousMonthEntries = getPreviousMonthEntries(allEntries);
      const buffer = createExcelBuffer(previousMonthEntries);

      const transporter = getMailTransporter();

      await transporter.sendMail({
        from: process.env.SMTP_USER,
        to: emailTo,
        subject: 'Merchant Monthly Report',
        text: 'Please find attached the previous month merchant report.',
        attachments: [
          {
            filename: 'merchant-monthly-report.xlsx',
            content: buffer,
          },
        ],
      });
    } catch {}
  },
  {
    timezone: 'Asia/Kolkata',
  }
);

/* ---------------- DATE RANGE EXCEL EXPORT ---------------- */

function parseDDMMYYYY(value) {
  const parts = String(value || '').split('-');

  if (parts.length !== 3) return null;

  const dd = Number(parts[0]);
  const mm = Number(parts[1]);
  const yyyy = Number(parts[2]);

  if (!dd || !mm || !yyyy) return null;

  return new Date(yyyy, mm - 1, dd);
}

app.get('/export/excel/date-range', (req, res) => {
  const { from, to } = req.query;

  const fromDate = parseDDMMYYYY(from);
  const toDate = parseDDMMYYYY(to);

  if (!fromDate || !toDate) {
    return res.status(400).json({
      error: 'Invalid date format. Use dd-mm-yyyy',
    });
  }

  toDate.setHours(23, 59, 59, 999);

  const entries = readJson(files.entries, []);

  const filtered = entries.filter((entry) => {
    const entryDate = new Date(entry.created_at || entry.createdAt || '');
    return entryDate >= fromDate && entryDate <= toDate;
  });

  const buffer = createExcelBuffer(filtered);

  res.setHeader(
    'Content-Disposition',
    `attachment; filename=merchant-data-${from}-to-${to}.xlsx`
  );

  res.setHeader(
    'Content-Type',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  );

  return res.send(buffer);
});

/* ---------------- 404 ---------------- */

app.use((req, res) => {
  res.status(404).json({
    error: 'Route not found',
    path: req.originalUrl,
  });
});

/* ---------------- SERVER ---------------- */

app.listen(PORT, '0.0.0.0');