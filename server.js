require('dotenv').config();

const express = require('express');
const cors = require('cors');
const XLSX = require('xlsx');
const multer = require('multer');
const fs = require('fs');
const { v2: cloudinary } = require('cloudinary');
const { createClient } = require('@supabase/supabase-js');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

function writeLog(level, message, data) {
  const time = new Date().toISOString();
  let output = `[${time}] [${level}] ${message}`;

  if (data !== undefined) {
    try {
      output += ` ${JSON.stringify(data)}`;
    } catch (err) {
      output += ` ${String(data)}`;
    }
  }

  process.stdout.write(`${output}\n`);
}

function logInfo(message, data) {
  writeLog('INFO', message, data);
}

function logError(message, data) {
  writeLog('ERROR', message, data);
}

if (!process.env.SUPABASE_URL || !process.env.SUPABASE_SERVICE_ROLE_KEY) {
  logError('Missing Supabase environment variables', {
    hasSupabaseUrl: !!process.env.SUPABASE_URL,
    hasServiceRoleKey: !!process.env.SUPABASE_SERVICE_ROLE_KEY,
  });
}

const supabase = createClient(
  process.env.SUPABASE_URL || '',
  process.env.SUPABASE_SERVICE_ROLE_KEY || '',
  {
    auth: {
      autoRefreshToken: false,
      persistSession: false,
      detectSessionInUrl: false,
    },
  }
);

cloudinary.config({
  cloud_name: process.env.CLOUDINARY_CLOUD_NAME,
  api_key: process.env.CLOUDINARY_API_KEY,
  api_secret: process.env.CLOUDINARY_API_SECRET,
});

const upload = multer({ dest: 'uploads/' });

function parseMessage(message) {
  const text = String(message || '');

  const name = text.match(/Name:\s*(.*)/i)?.[1]?.trim() || '';
  const product = text.match(/Product:\s*(.*)/i)?.[1]?.trim() || '';
  const quantity = text.match(/Quantity:\s*(.*)/i)?.[1]?.trim() || '';
  const time = text.match(/Time:\s*(.*)/i)?.[1]?.trim() || '';

  return { name, product, quantity, time };
}

function extractQuantityNumber(quantity) {
  const match = String(quantity || '').match(/[\d.]+/);
  return match ? parseFloat(match[0]) : 0;
}

function buildStats(entriesList) {
  const today = getTodayDateString();

  const todayEntries = (entriesList || []).filter(
    (entry) => String(entry.created_at || '').slice(0, 10) === today
  );

  const totalQuantity = todayEntries.reduce((sum, entry) => {
    return sum + extractQuantityNumber(entry.quantity || '');
  }, 0);

  const productSummary = {};

  todayEntries.forEach((entry) => {
    const productName = entry.product || 'Unknown';
    if (!productSummary[productName]) {
      productSummary[productName] = 0;
    }
    productSummary[productName] += extractQuantityNumber(entry.quantity || '');
  });

  return {
    trackedToday: todayEntries.length,
    totalLoads: todayEntries.length,
    totalQuantity,
    productSummary,
  };
}

function cleanPhone(value) {
  return String(value || '').trim();
}

function normalizePhone(value) {
  return String(value || '')
    .replace(/^whatsapp:/i, '')
    .replace(/\s+/g, '')
    .trim();
}

function isValidTimeString(value) {
  return /^([01]\d|2[0-3]):([0-5]\d)$/.test(String(value || ''));
}

function timeToMinutes(value) {
  const [hours, minutes] = String(value).split(':').map(Number);
  return hours * 60 + minutes;
}

function getIndiaCurrentTimeHHMM() {
  const formatter = new Intl.DateTimeFormat('en-GB', {
    timeZone: 'Asia/Kolkata',
    hour: '2-digit',
    minute: '2-digit',
    hour12: false,
  });

  const parts = formatter.formatToParts(new Date());
  const hour = parts.find((p) => p.type === 'hour')?.value || '00';
  const minute = parts.find((p) => p.type === 'minute')?.value || '00';

  return `${hour}:${minute}`;
}

function isWithinTrackingWindow(currentHHMM, startHHMM, endHHMM) {
  if (!isValidTimeString(startHHMM) || !isValidTimeString(endHHMM)) {
    return true;
  }

  const current = timeToMinutes(currentHHMM);
  const start = timeToMinutes(startHHMM);
  const end = timeToMinutes(endHHMM);

  if (start <= end) {
    return current >= start && current <= end;
  }

  return current >= start || current <= end;
}

async function getTrackedMerchantConfig(senderNumber) {
  const normalizedSender = normalizePhone(senderNumber);

  const { data, error } = await supabase
    .from('tracked_merchants')
    .select('*')
    .eq('merchant_number', normalizedSender)
    .maybeSingle();

  if (error) throw error;
  return data || null;
}

function formatDateTime(value) {
  if (!value) return '';
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return String(value);

  return date.toLocaleString('en-IN', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
  });
}

function getTodayDateString() {
  const now = new Date();

  const year = now.toLocaleString("en-US", {
    timeZone: "Asia/Kolkata",
    year: "numeric",
  });

  const month = now.toLocaleString("en-US", {
    timeZone: "Asia/Kolkata",
    month: "2-digit",
  });

  const day = now.toLocaleString("en-US", {
    timeZone: "Asia/Kolkata",
    day: "2-digit",
  });

  return `${year}-${month}-${day}`;
}

function getCurrentMonthPrefix() {
  const now = new Date();

  const year = now.toLocaleString("en-US", {
    timeZone: "Asia/Kolkata",
    year: "numeric",
  });

  const month = now.toLocaleString("en-US", {
    timeZone: "Asia/Kolkata",
    month: "2-digit",
  });

  return `${year}-${month}`;
}

function buildEntryExportRows(entries) {
  return (entries || []).map((entry, index) => ({
    'S.No': index + 1,
    'Merchant Name': entry.merchant_name || '',
    'Sender Number': entry.sender || '',
    'Receiver Number': entry.receiver || '',
    Product: entry.product || '',
    Quantity: entry.quantity || '',
    Time: entry.time || '',
    'Created At': formatDateTime(entry.created_at),
  }));
}

async function getFilteredEntries(filterType) {
  const { data, error } = await supabase
    .from('entries')
    .select('*')
    .order('created_at', { ascending: false });

  if (error) throw error;

  const allEntries = data || [];

  if (filterType === 'daily') {
    const today = getTodayDateString();
    return allEntries.filter(
      (entry) => String(entry.created_at || '').slice(0, 10) === today
    );
  }

  if (filterType === 'monthly') {
    const monthPrefix = getCurrentMonthPrefix();
    return allEntries.filter(
      (entry) => String(entry.created_at || '').slice(0, 7) === monthPrefix
    );
  }

  return allEntries;
}

function sendEntriesWorkbook(res, entries, reportTitle, fileName) {
  const workbook = XLSX.utils.book_new();
  const rows = buildEntryExportRows(entries);
  const sheet = XLSX.utils.json_to_sheet(rows);

  XLSX.utils.book_append_sheet(workbook, sheet, reportTitle);

  const buffer = XLSX.write(workbook, {
    type: 'buffer',
    bookType: 'xlsx',
  });

  res.setHeader(
    'Content-Disposition',
    `attachment; filename="${fileName}"`
  );
  res.setHeader(
    'Content-Type',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  );

  return res.send(buffer);
}

/* -------------------- BASIC -------------------- */

app.get('/', (req, res) => {
  res.send('Backend running with Supabase');
});

app.get('/health', (req, res) => {
  res.json({
    ok: true,
    message: 'Server is running',
    port: PORT,
    file: 'tracked-merchant-and-business-numbers',
    indiaTime: getIndiaCurrentTimeHHMM(),
    time: new Date().toISOString(),
  });
});

app.get('/routes-check', (req, res) => {
  res.json({
    message: 'These routes are active',
    routes: [
      '/',
      '/health',
      '/routes-check',
      '/entries',
      '/stats',
      '/products',
      '/profile',
      '/numbers',
      '/numbers/:id/status',
      '/numbers/:id',
      '/tracked-merchants',
      '/tracked-merchants/:id/status',
      '/tracked-merchants/:id/schedule',
      '/tracked-merchants/:id',
      '/webhook',
      '/upload/image',
      '/export/excel',
      '/export/excel/daily',
      '/export/excel/monthly',
      '/export/excel/total',
    ],
  });
});

/* -------------------- ENTRIES -------------------- */

app.get('/entries', async (req, res) => {
  try {
    const { data, error } = await supabase
      .from('entries')
      .select('*')
      .order('created_at', { ascending: false });

    if (error) {
      logError('Entries fetch error', error);
      return res.status(500).json({ error: 'Failed to load entries', details: error.message });
    }

    return res.json(data || []);
  } catch (error) {
    logError('Entries fetch exception', error);
    return res.status(500).json({ error: 'Failed to load entries', details: error.message });
  }
});

app.get('/stats', async (req, res) => {
  try {
    const { data, error } = await supabase
      .from('entries')
      .select('*')
      .order('created_at', { ascending: false });

    if (error) {
      logError('Stats fetch error', error);
      return res.status(500).json({ error: 'Failed to load stats', details: error.message });
    }

    return res.json(buildStats(data || []));
  } catch (error) {
    logError('Stats fetch exception', error);
    return res.status(500).json({ error: 'Failed to load stats', details: error.message });
  }
});

/* -------------------- BUSINESS NUMBERS -------------------- */

app.get('/numbers', async (req, res) => {
  try {
    const { data, error } = await supabase
      .from('business_numbers')
      .select('*')
      .order('created_at', { ascending: false });

    if (error) {
      logError('GET /numbers error', error);
      return res.status(500).json({
        error: 'Failed to fetch business numbers',
        details: error.message,
      });
    }

    return res.json(data || []);
  } catch (error) {
    logError('GET /numbers exception', error);
    return res.status(500).json({
      error: 'Failed to fetch business numbers',
      details: error.message,
    });
  }
});

app.post('/numbers', async (req, res) => {
  try {
    const { label, phone_number } = req.body;

    if (!label || !phone_number) {
      return res.status(400).json({
        error: 'Label and phone number are required',
      });
    }

    const normalizedPhoneNumber = normalizePhone(phone_number);

    const { data, error } = await supabase
      .from('business_numbers')
      .insert([
        {
          label,
          phone_number: normalizedPhoneNumber,
          status: 'active',
        },
      ])
      .select();

    if (error) {
      logError('POST /numbers error', error);
      return res.status(500).json({
        error: 'Failed to add business number',
        details: error.message,
      });
    }

    return res.json(data[0]);
  } catch (error) {
    logError('POST /numbers exception', error);
    return res.status(500).json({
      error: 'Failed to add business number',
      details: error.message,
    });
  }
});

app.put('/numbers/:id/status', async (req, res) => {
  try {
    const { id } = req.params;
    const { status } = req.body;

    if (!id || !id.trim()) {
      return res.status(400).json({ error: 'Business number id is required' });
    }

    if (status !== 'active' && status !== 'inactive') {
      return res.status(400).json({
        error: 'Status must be either active or inactive',
      });
    }

    const { data, error } = await supabase
      .from('business_numbers')
      .update({ status })
      .eq('id', id.trim())
      .select();

    if (error) {
      logError('PUT /numbers/:id/status error', error);
      return res.status(500).json({
        error: 'Failed to update business number status',
        details: error.message,
      });
    }

    if (!data || data.length === 0) {
      return res.status(404).json({ error: 'Business number not found' });
    }

    return res.json({
      message: `Business number ${status} successfully`,
      number: data[0],
    });
  } catch (error) {
    logError('PUT /numbers/:id/status exception', error);
    return res.status(500).json({
      error: 'Failed to update business number status',
      details: error.message,
    });
  }
});

app.delete('/numbers/:id', async (req, res) => {
  try {
    const { id } = req.params;

    if (!id || !id.trim()) {
      return res.status(400).json({ error: 'Valid business number id is required' });
    }

    const { data, error } = await supabase
      .from('business_numbers')
      .delete()
      .eq('id', id.trim())
      .select();

    if (error) {
      logError('DELETE /numbers error', error);
      return res.status(500).json({
        error: 'Failed to delete business number',
        details: error.message,
      });
    }

    if (!data || data.length === 0) {
      return res.status(404).json({ error: 'Business number not found' });
    }

    return res.json({
      message: 'Business number deleted successfully',
      number: data[0],
    });
  } catch (error) {
    logError('DELETE /numbers exception', error);
    return res.status(500).json({
      error: 'Failed to delete business number',
      details: error.message,
    });
  }
});

/* -------------------- TRACKED MERCHANT WEBHOOK -------------------- */

app.post('/webhook', async (req, res) => {
  const message = req.body.Body || '';
  const sender = cleanPhone(req.body.From);
  const receiver = cleanPhone(req.body.To);

  const normalizedSender = normalizePhone(sender);
  const normalizedReceiver = normalizePhone(receiver);

  const data = parseMessage(message);
  res.set('Content-Type', 'text/xml');

  try {
    const merchantConfig = await getTrackedMerchantConfig(normalizedSender);

    if (!merchantConfig) {
      logInfo('Webhook ignored because sender merchant is not configured', {
        sender: normalizedSender,
        receiver: normalizedReceiver,
      });

      return res.status(200).send(
`<?xml version="1.0" encoding="UTF-8"?>
<Response>
  <Message>Your number is not configured for tracking. Please contact admin.</Message>
</Response>`
      );
    }

    if (merchantConfig.status !== 'active') {
      logInfo('Webhook ignored because merchant is inactive', {
        sender: normalizedSender,
        merchant: merchantConfig.merchant_name,
      });

      return res.status(200).send(
`<?xml version="1.0" encoding="UTF-8"?>
<Response>
  <Message>Tracking is inactive for your number. Please contact admin.</Message>
</Response>`
      );
    }

    const startTime = merchantConfig.tracking_start_time || '00:00';
    const endTime = merchantConfig.tracking_end_time || '23:59';
    const currentIndiaTime = getIndiaCurrentTimeHHMM();

    const allowedTime = isWithinTrackingWindow(
      currentIndiaTime,
      startTime,
      endTime
    );

    if (!allowedTime) {
      logInfo('Webhook ignored because merchant time period completed', {
        sender: normalizedSender,
        merchant: merchantConfig.merchant_name,
        currentIndiaTime,
        startTime,
        endTime,
      });

      return res.status(200).send(
`<?xml version="1.0" encoding="UTF-8"?>
<Response>
  <Message>Your time period was completed. Please contact admin.</Message>
</Response>`
      );
    }

    if (data.name || data.product || data.quantity || data.time) {
      const payload = {
  sender: normalizedSender,
  receiver: normalizedReceiver,
  merchant_name: merchantConfig.merchant_name || data.name,
  product: data.product,
  quantity: data.quantity,
  time: data.time,
  created_at: new Date().toLocaleString("sv-SE", {
    timeZone: "Asia/Kolkata"
  }).replace(" ", "T"),
};

      const { error } = await supabase.from('entries').insert([payload]);

      if (error) {
        logError('Webhook insert error', error);

        return res.status(200).send(
`<?xml version="1.0" encoding="UTF-8"?>
<Response>
  <Message>Failed to save entry</Message>
</Response>`
        );
      }

      logInfo('Webhook entry stored for tracked merchant', {
        ...payload,
        trackedWindow: `${startTime}-${endTime}`,
        currentIndiaTime,
      });

      return res.status(200).send(
`<?xml version="1.0" encoding="UTF-8"?>
<Response>
  <Message>Entry recorded successfully.</Message>
</Response>`
      );
    }

    return res.status(200).send(
`<?xml version="1.0" encoding="UTF-8"?>
<Response>
  <Message>Invalid format. Please send: Name, Product, Quantity, Time.</Message>
</Response>`
    );
  } catch (error) {
    logError('Webhook exception', error);

    return res.status(200).send(
`<?xml version="1.0" encoding="UTF-8"?>
<Response>
  <Message>Webhook error</Message>
</Response>`
    );
  }
});

/* -------------------- TRACKED MERCHANTS -------------------- */

app.get('/tracked-merchants', async (req, res) => {
  try {
    const { data, error } = await supabase
      .from('tracked_merchants')
      .select('*')
      .order('created_at', { ascending: false });

    if (error) {
      logError('GET /tracked-merchants error', error);
      return res.status(500).json({
        error: 'Failed to fetch tracked merchants',
        details: error.message,
      });
    }

    return res.json(data || []);
  } catch (error) {
    logError('GET /tracked-merchants exception', error);
    return res.status(500).json({
      error: 'Failed to fetch tracked merchants',
      details: error.message,
    });
  }
});

app.post('/tracked-merchants', async (req, res) => {
  try {
    const {
      merchant_name,
      merchant_number,
      tracking_start_time,
      tracking_end_time,
    } = req.body;

    if (!merchant_name || !merchant_number) {
      return res.status(400).json({
        error: 'Merchant name and merchant number are required',
      });
    }

    const normalizedMerchantNumber = normalizePhone(merchant_number);
    const startTime = tracking_start_time || '00:00';
    const endTime = tracking_end_time || '23:59';

    if (!isValidTimeString(startTime) || !isValidTimeString(endTime)) {
      return res.status(400).json({
        error: 'Tracking start and end time must be in HH:MM format',
      });
    }

    const { data, error } = await supabase
      .from('tracked_merchants')
      .insert([
        {
          merchant_name,
          merchant_number: normalizedMerchantNumber,
          status: 'active',
          tracking_start_time: startTime,
          tracking_end_time: endTime,
        },
      ])
      .select();

    if (error) {
      logError('POST /tracked-merchants error', error);
      return res.status(500).json({
        error: 'Failed to add tracked merchant',
        details: error.message,
      });
    }

    return res.json(data[0]);
  } catch (error) {
    logError('POST /tracked-merchants exception', error);
    return res.status(500).json({
      error: 'Failed to add tracked merchant',
      details: error.message,
    });
  }
});

app.put('/tracked-merchants/:id/status', async (req, res) => {
  try {
    const { id } = req.params;
    const { status } = req.body;

    if (!id || !id.trim()) {
      return res.status(400).json({ error: 'Merchant id is required' });
    }

    if (status !== 'active' && status !== 'inactive') {
      return res.status(400).json({
        error: 'Status must be either active or inactive',
      });
    }

    const { data, error } = await supabase
      .from('tracked_merchants')
      .update({ status })
      .eq('id', id.trim())
      .select();

    if (error) {
      logError('PUT /tracked-merchants/:id/status error', error);
      return res.status(500).json({
        error: 'Failed to update merchant status',
        details: error.message,
      });
    }

    if (!data || data.length === 0) {
      return res.status(404).json({ error: 'Merchant not found' });
    }

    return res.json({
      message: `Merchant ${status} successfully`,
      merchant: data[0],
    });
  } catch (error) {
    logError('PUT /tracked-merchants/:id/status exception', error);
    return res.status(500).json({
      error: 'Failed to update merchant status',
      details: error.message,
    });
  }
});

app.put('/tracked-merchants/:id/schedule', async (req, res) => {
  try {
    const { id } = req.params;
    const { tracking_start_time, tracking_end_time } = req.body;

    if (!id || !id.trim()) {
      return res.status(400).json({ error: 'Merchant id is required' });
    }

    if (
      !isValidTimeString(tracking_start_time) ||
      !isValidTimeString(tracking_end_time)
    ) {
      return res.status(400).json({
        error: 'Tracking start and end time must be in HH:MM format',
      });
    }

    const { data, error } = await supabase
      .from('tracked_merchants')
      .update({
        tracking_start_time,
        tracking_end_time,
      })
      .eq('id', id.trim())
      .select();

    if (error) {
      logError('PUT /tracked-merchants/:id/schedule error', error);
      return res.status(500).json({
        error: 'Failed to update merchant schedule',
        details: error.message,
      });
    }

    if (!data || data.length === 0) {
      return res.status(404).json({ error: 'Merchant not found' });
    }

    return res.json({
      message: 'Merchant schedule updated successfully',
      merchant: data[0],
    });
  } catch (error) {
    logError('PUT /tracked-merchants/:id/schedule exception', error);
    return res.status(500).json({
      error: 'Failed to update merchant schedule',
      details: error.message,
    });
  }
});

app.delete('/tracked-merchants/:id', async (req, res) => {
  try {
    const { id } = req.params;

    if (!id || !id.trim()) {
      return res.status(400).json({ error: 'Valid merchant id is required' });
    }

    const { data, error } = await supabase
      .from('tracked_merchants')
      .delete()
      .eq('id', id.trim())
      .select();

    if (error) {
      logError('DELETE /tracked-merchants error', error);
      return res.status(500).json({
        error: 'Failed to delete tracked merchant',
        details: error.message,
      });
    }

    if (!data || data.length === 0) {
      return res.status(404).json({ error: 'Merchant not found' });
    }

    return res.json({
      message: 'Tracked merchant deleted successfully',
      merchant: data[0],
    });
  } catch (error) {
    logError('DELETE /tracked-merchants exception', error);
    return res.status(500).json({
      error: 'Failed to delete tracked merchant',
      details: error.message,
    });
  }
});

/* -------------------- PRODUCTS -------------------- */

app.get('/products', async (req, res) => {
  try {
    const { data, error } = await supabase
      .from('products')
      .select('*')
      .order('created_at', { ascending: false });

    if (error) {
      logError('Products fetch error', error);
      return res.status(500).json({ error: 'Failed to load products', details: error.message });
    }

    return res.json(data || []);
  } catch (error) {
    logError('Products fetch exception', error);
    return res.status(500).json({ error: 'Failed to load products', details: error.message });
  }
});

app.post('/products', async (req, res) => {
  try {
    const { name, title, image, product_number } = req.body;

    if (!name || !title) {
      return res.status(400).json({ error: 'Product name and title are required' });
    }

    const { data, error } = await supabase
      .from('products')
      .insert([
        {
          name,
          title,
          image: image || '',
          product_number: product_number ?? null,
        },
      ])
      .select();

    if (error) {
      logError('Insert product error', error);
      return res.status(500).json({ error: 'Failed to save product', details: error.message });
    }

    return res.json(data[0]);
  } catch (error) {
    logError('Insert product exception', error);
    return res.status(500).json({ error: 'Failed to save product', details: error.message });
  }
});

app.put('/products/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const { name, title, image, product_number } = req.body;

    const { data, error } = await supabase
      .from('products')
      .update({
        name,
        title,
        image,
        product_number,
      })
      .eq('id', id)
      .select();

    if (error) {
      logError('Update product error', error);
      return res.status(500).json({ error: 'Failed to update product', details: error.message });
    }

    if (!data || data.length === 0) {
      return res.status(404).json({ error: 'Product not found' });
    }

    return res.json(data[0]);
  } catch (error) {
    logError('Update product exception', error);
    return res.status(500).json({ error: 'Failed to update product', details: error.message });
  }
});

app.delete('/products/:id', async (req, res) => {
  try {
    const { id } = req.params;

    const { error } = await supabase
      .from('products')
      .delete()
      .eq('id', id);

    if (error) {
      logError('Delete product error', error);
      return res.status(500).json({ error: 'Failed to delete product', details: error.message });
    }

    return res.json({ message: 'Product deleted successfully' });
  } catch (error) {
    logError('Delete product exception', error);
    return res.status(500).json({ error: 'Failed to delete product', details: error.message });
  }
});

/* -------------------- PROFILE -------------------- */

app.get('/profile', async (req, res) => {
  try {
    const { data, error } = await supabase
      .from('profiles')
      .select('*')
      .limit(1);

    if (error) {
      logError('GET /profile error', error);
      return res.status(500).json({ error: 'Failed to load profile', details: error.message });
    }

    return res.json(data[0] || null);
  } catch (error) {
    logError('GET /profile exception', error);
    return res.status(500).json({ error: 'Failed to load profile', details: error.message });
  }
});

app.put('/profile', async (req, res) => {
  try {
    const { name, phone, email } = req.body;

    const { data: existing, error: fetchError } = await supabase
      .from('profiles')
      .select('*')
      .limit(1);

    if (fetchError) {
      logError('PUT /profile fetch error', fetchError);
      return res.status(500).json({ error: 'Failed to load profile', details: fetchError.message });
    }

    if (!existing || existing.length === 0) {
      const { data, error } = await supabase
        .from('profiles')
        .insert([{ name, phone, email, password: '' }])
        .select();

      if (error) {
        logError('PUT /profile insert error', error);
        return res.status(500).json({ error: 'Failed to save profile', details: error.message });
      }

      return res.json(data[0]);
    }

    const profileId = existing[0].id;

    const { data, error } = await supabase
      .from('profiles')
      .update({ name, phone, email })
      .eq('id', profileId)
      .select();

    if (error) {
      logError('PUT /profile update error', error);
      return res.status(500).json({ error: 'Failed to update profile', details: error.message });
    }

    return res.json(data[0]);
  } catch (error) {
    logError('PUT /profile exception', error);
    return res.status(500).json({ error: 'Failed to update profile', details: error.message });
  }
});

/* -------------------- IMAGE UPLOAD -------------------- */

app.post('/upload/image', upload.single('image'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No image file uploaded' });
    }

    const result = await cloudinary.uploader.upload(req.file.path, {
      folder: 'merchantdispatch/products',
    });

    try {
      fs.unlinkSync(req.file.path);
    } catch (unlinkError) {
      logError('Temporary file delete error', unlinkError);
    }

    return res.json({
      imageUrl: result.secure_url,
      publicId: result.public_id,
    });
  } catch (error) {
    logError('Cloudinary upload error', error);
    return res.status(500).json({ error: 'Failed to upload image', details: error.message });
  }
});

/* -------------------- EXPORT -------------------- */

app.get('/export/excel', async (req, res) => {
  try {
    const [
      { data: entriesData, error: entriesError },
      { data: productsData, error: productsError },
    ] = await Promise.all([
      supabase.from('entries').select('*').order('created_at', { ascending: false }),
      supabase.from('products').select('*').order('created_at', { ascending: false }),
    ]);

    if (entriesError || productsError) {
      logError('Export data fetch error', {
        entriesError,
        productsError,
      });
      return res.status(500).json({
        error: 'Failed to load export data',
        details: entriesError?.message || productsError?.message || 'Unknown export error',
      });
    }

    const workbook = XLSX.utils.book_new();
    const stats = buildStats(entriesData || []);

    const summaryData = [
      { label: 'Tracked Today', value: stats.trackedToday || 0 },
      { label: 'Total Loads', value: stats.totalLoads || 0 },
      { label: 'Total Quantity', value: stats.totalQuantity || 0 },
      { label: 'Total Products', value: (productsData || []).length || 0 },
      { label: 'Exported At', value: new Date().toISOString() },
    ];

    const summarySheet = XLSX.utils.json_to_sheet(summaryData);
    const entriesSheet = XLSX.utils.json_to_sheet(entriesData || []);
    const productsSheet = XLSX.utils.json_to_sheet(productsData || []);

    XLSX.utils.book_append_sheet(workbook, summarySheet, 'Summary');
    XLSX.utils.book_append_sheet(workbook, entriesSheet, 'Entries');
    XLSX.utils.book_append_sheet(workbook, productsSheet, 'Products');

    const buffer = XLSX.write(workbook, {
      type: 'buffer',
      bookType: 'xlsx',
    });

    res.setHeader(
      'Content-Disposition',
      'attachment; filename="merchant-dashboard-data.xlsx"'
    );
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );

    return res.send(buffer);
  } catch (error) {
    logError('Excel export error', error);
    return res.status(500).json({ error: 'Failed to export Excel file', details: error.message });
  }
});

app.get('/export/excel/daily', async (req, res) => {
  try {
    const entries = await getFilteredEntries('daily');
    return sendEntriesWorkbook(res, entries, 'Daily Data', 'merchant-daily-data.xlsx');
  } catch (error) {
    logError('Daily export error', error);
    return res.status(500).json({
      error: 'Failed to export daily data',
      details: error.message,
    });
  }
});

app.get('/export/excel/monthly', async (req, res) => {
  try {
    const entries = await getFilteredEntries('monthly');
    return sendEntriesWorkbook(res, entries, 'Monthly Data', 'merchant-monthly-data.xlsx');
  } catch (error) {
    logError('Monthly export error', error);
    return res.status(500).json({
      error: 'Failed to export monthly data',
      details: error.message,
    });
  }
});

app.get('/export/excel/total', async (req, res) => {
  try {
    const entries = await getFilteredEntries('total');
    return sendEntriesWorkbook(res, entries, 'Total Data', 'merchant-total-data.xlsx');
  } catch (error) {
    logError('Total export error', error);
    return res.status(500).json({
      error: 'Failed to export total data',
      details: error.message,
    });
  }
});

/* -------------------- 404 -------------------- */

app.use((req, res) => {
  res.status(404).json({
    error: 'Route not found',
    method: req.method,
    path: req.originalUrl,
    hint: 'Check /routes-check to confirm active routes',
  });
});

app.listen(PORT, '0.0.0.0', () => {
  logInfo(`Server running on port ${PORT} with Supabase`);
  logInfo('Active test URLs', {
    root: `http://localhost:${PORT}/`,
    health: `http://localhost:${PORT}/health`,
    routesCheck: `http://localhost:${PORT}/routes-check`,
    businessNumbers: `http://localhost:${PORT}/numbers`,
    trackedMerchants: `http://localhost:${PORT}/tracked-merchants`,
    webhook: `http://localhost:${PORT}/webhook`,
  });
});