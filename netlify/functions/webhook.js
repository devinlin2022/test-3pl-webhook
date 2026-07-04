const { google } = require('googleapis');
const busboy = require('busboy');
const xlsx = require('xlsx');

// Note: Ensure you set these in your Netlify Environment Variables
// GOOGLE_SERVICE_ACCOUNT_EMAIL
// GOOGLE_PRIVATE_KEY (replace \n with actual newlines in Netlify UI or handle parsing below)
// GOOGLE_SPREADSHEET_ID
// GOOGLE_SHEET_NAME (e.g., 'Sheet1')

const getGoogleSheetsClient = () => {
  const privateKey = process.env.GOOGLE_PRIVATE_KEY 
    ? process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n') 
    : undefined;

  if (!process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || !privateKey) {
    throw new Error('Google Service Account credentials missing in environment variables.');
  }

  const auth = new google.auth.JWT(
    process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
    null,
    privateKey,
    ['https://www.googleapis.com/auth/spreadsheets']
  );

  return google.sheets({ version: 'v4', auth });
};

const appendToGoogleSheet = async (dataRows) => {
  if (dataRows.length === 0) return;

  const sheets = getGoogleSheetsClient();
  const spreadsheetId = process.env.GOOGLE_SPREADSHEET_ID;
  const sheetName = process.env.GOOGLE_SHEET_NAME || 'Sheet1';

  // Format data as an array of arrays in the correct column order:
  // Container, MBL, Status, Forwarder, Vessel ETD, Vessel ATD, Initial Vessel ETA, Vessel ETA
  const values = dataRows.map(row => [
    row['Container'] || '',
    row['MBL'] || '',
    row['Status'] || '',
    row['Forwarder'] || '',
    row['Vessel ETD'] || '',
    row['Vessel ATD'] || '',
    row['Initial Vessel ETA'] || '',
    row['Vessel ETA'] || ''
  ]);

  await sheets.spreadsheets.values.append({
    spreadsheetId,
    range: `${sheetName}!A:H`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values },
  });
};

const parseMultipart = (event) => {
  return new Promise((resolve, reject) => {
    const bb = busboy({ headers: event.headers });
    let fileData = null;
    let fileName = '';

    bb.on('file', (name, file, info) => {
      fileName = info.filename;
      const chunks = [];
      file.on('data', (data) => {
        chunks.push(data);
      });
      file.on('end', () => {
        fileData = Buffer.concat(chunks);
      });
    });

    bb.on('finish', () => {
      if (!fileData) {
        return reject(new Error('No file uploaded'));
      }
      resolve({ fileData, fileName });
    });

    bb.on('error', (err) => {
      reject(err);
    });

    bb.end(Buffer.from(event.body, event.isBase64Encoded ? 'base64' : 'utf-8'));
  });
};

const parseExcelFile = (buffer) => {
  const workbook = xlsx.read(buffer, { type: 'buffer' });
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  // Convert sheet to JSON array
  const json = xlsx.utils.sheet_to_json(worksheet);
  return json;
};

exports.handler = async (event, context) => {
  // Only allow POST
  if (event.httpMethod !== 'POST') {
    return {
      statusCode: 405,
      body: JSON.stringify({ error: 'Method Not Allowed. Use POST.' }),
    };
  }

  try {
    let shipmentsData = [];
    const contentType = event.headers['content-type'] || event.headers['Content-Type'] || '';

    if (contentType.includes('multipart/form-data')) {
      const { fileData } = await parseMultipart(event);
      shipmentsData = parseExcelFile(fileData);
    } else {
      // Default to trying to parse as JSON if it's not multipart
      try {
        const bodyStr = event.isBase64Encoded 
          ? Buffer.from(event.body, 'base64').toString('utf-8') 
          : event.body;
        
        const payload = JSON.parse(bodyStr);
        shipmentsData = Array.isArray(payload) ? payload : [payload];
      } catch (err) {
        return {
          statusCode: 415,
          body: JSON.stringify({ error: 'Unsupported Media Type or Invalid JSON. Please send application/json or multipart/form-data.' }),
        };
      }
    }

    if (shipmentsData.length > 1000) {
      return {
        statusCode: 413,
        body: JSON.stringify({ error: 'Payload too large. Maximum 1000 records per batch.' }),
      };
    }

    await appendToGoogleSheet(shipmentsData);

    return {
      statusCode: 200,
      body: JSON.stringify({ message: `Successfully processed ${shipmentsData.length} records.` }),
    };

  } catch (error) {
    console.error('Error processing webhook:', error);
    return {
      statusCode: 500,
      body: JSON.stringify({ error: 'Internal Server Error', details: error.message }),
    };
  }
};
