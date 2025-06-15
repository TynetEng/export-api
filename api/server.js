require('dotenv').config();
const express = require('express');
const cors = require('cors');
const axios = require('axios');
const bodyParser = require('body-parser');
const nodemailer = require('nodemailer');
const puppeteer = require('puppeteer');

const app = express();
app.use(cors());
app.use(bodyParser.json());

// Get Microsoft Graph access token
const getAccessToken = async () => {
  const { CLIENT_ID, CLIENT_SECRET, TENANT_ID } = process.env;
  const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;

  const params = new URLSearchParams();
  params.append('client_id', CLIENT_ID);
  params.append('client_secret', CLIENT_SECRET);
  params.append('grant_type', 'client_credentials');
  params.append('scope', 'https://graph.microsoft.com/.default');

  const response = await axios.post(url, params, {
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
  });
  return response.data.access_token;
};

// Get Graph site ID
const getGraphSiteId = async (token) => {
  const { SHAREPOINT_SITE_HOST, SHAREPOINT_SITE_PATH } = process.env;

  const response = await axios.get(
    `https://graph.microsoft.com/v1.0/sites/${SHAREPOINT_SITE_HOST}:/sites/${SHAREPOINT_SITE_PATH}`,
    {
      headers: { Authorization: `Bearer ${token}` },
    }
  );
  return response.data.id;
};

// Get SharePoint list ID by name
const getListId = async (token, siteId, listName) => {
  const response = await axios.get(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists`,
    {
      headers: { Authorization: `Bearer ${token}` },
    }
  );

  const list = response.data.value.find((l) => l.displayName === listName);
  if (!list) throw new Error(`List '${listName}' not found`);
  return list.id;
};

// Get specific item by ID
const getItemById = async (token, siteId, listId, itemId) => {
  const response = await axios.get(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${itemId}?expand=fields`,
    {
      headers: { Authorization: `Bearer ${token}` },
    }
  );
  return response.data;
};

// Route to fetch a SharePoint item by ID
app.get('/api/item/:id', async (req, res) => {
  try {
    const token = await getAccessToken();
    const siteId = await getGraphSiteId(token);
    const listId = await getListId(token, siteId, process.env.SHAREPOINT_LIST_NAME);
    const item = await getItemById(token, siteId, listId, req.params.id);

    res.json(item);
  } catch (err) {
    console.error('Error fetching SharePoint item:', err.response?.data || err.message);
    res.status(500).json({
      error: 'Failed to fetch SharePoint item',
      details: err.response?.data || err.message,
    });
  }
});

// Debug route: List all SharePoint lists
app.get('/api/lists', async (req, res) => {
  try {
    const token = await getAccessToken();
    const siteId = await getGraphSiteId(token);
    const response = await axios.get(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/lists`,
      {
        headers: { Authorization: `Bearer ${token}` },
      }
    );

    const lists = response.data.value.map((l) => ({
      displayName: l.displayName,
      id: l.id,
    }));
    res.json(lists);
  } catch (err) {
    console.error('Error fetching SharePoint lists:', err.response?.data || err.message);
    res.status(500).json({
      error: 'Failed to fetch SharePoint lists',
      details: err.response?.data || err.message,
    });
  }
});

// Convert HTML to PDF using Puppeteer
async function generatePdfFromHtml(htmlContent) {
  const browser = await puppeteer.launch({ headless: 'new' });
  const page = await browser.newPage();
  await page.setContent(htmlContent, { waitUntil: 'networkidle0' });
  const pdfBuffer = await page.pdf({ format: 'A4' });
  await browser.close();
  return pdfBuffer;
}

// Submit shipping form and send email
app.post('/api/submit-shipping', async (req, res) => {
  try {
    const payload = req.body;

    const containerRows = payload.containers.map((c, i) => `
      <tr>
        <td>${i + 1}</td>
        <td>${c.containerNumber}</td>
        <td>${c.description}</td>
        <td>${c.quantity}</td>
        <td>${c.value}</td>
        <td>${c.hsCode}</td>
        <td>${c.weight}</td>
      </tr>
    `).join('');

    const htmlContent = `
      <h2 style="color:#1a73e8; font-family:Arial;">Shipping Instruction</h2>
      <p><strong>Submitted By:</strong> ${payload.user?.name} (${payload.user?.email})</p>
      <p><strong>Carrier Reference:</strong> ${payload.carrierReference}</p>
      <hr />
      <h3>Billing Party</h3>
      <p>${payload.billingParty?.name}<br/>
      ${payload.billingParty?.address1}, ${payload.billingParty?.address2}<br/>
      ${payload.billingParty?.city}, ${payload.billingParty?.country} - ${payload.billingParty?.postcode}<br/>
      Email: ${payload.billingParty?.email}, Phone: ${payload.billingParty?.phone}</p>

      <h3>Shipper</h3>
      <p>${payload.shipper?.name}<br/>
      ${payload.shipper?.address1}, ${payload.shipper?.address2}<br/>
      ${payload.shipper?.city}, ${payload.shipper?.country} - ${payload.shipper?.postcode}<br/>
      Email: ${payload.shipper?.email}, Phone: ${payload.shipper?.phone}</p>

      <h3>Consignee</h3>
      <p>${payload.consignee?.name}<br/>
      ${payload.consignee?.address1}, ${payload.consignee?.address2}<br/>
      ${payload.consignee?.city}, ${payload.consignee?.country} - ${payload.consignee?.postcode}<br/>
      Email: ${payload.consignee?.email}, Phone: ${payload.consignee?.phone}</p>

      <h3>Shipment Details</h3>
      <p><strong>Value:</strong> ${payload.shipmentValue}</p>
      <p><strong>Notes:</strong> ${payload.notes}</p>

      <h3>Containers</h3>
      <table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; width: 100%; font-size: 12px;">
        <thead>
          <tr>
            <th>#</th>
            <th>Container Number</th>
            <th>Description</th>
            <th>Quantity</th>
            <th>Value</th>
            <th>HS Code</th>
            <th>Weight</th>
          </tr>
        </thead>
        <tbody>
          ${containerRows}
        </tbody>
      </table>
    `;

    const pdfBuffer = await generatePdfFromHtml(htmlContent);

    // Nodemailer config
    const transporter = nodemailer.createTransport({
      service: 'gmail',
      auth: {
        user: process.env.SMTP_USER,
        pass: process.env.SMTP_PASS,
      },
    });

    await transporter.sendMail({
      from: `"Shipping Desk" <${process.env.SMTP_USER}>`,
      to: payload.user?.email || process.env.FALLBACK_RECIPIENT,
      subject: 'New Shipping Instruction Submission',
      html: htmlContent,
      attachments: [
        {
          filename: 'shipping-instruction.pdf',
          content: pdfBuffer,
        },
      ],
    });

    res.json({ message: 'Form submitted and email sent.' });
  } catch (err) {
    console.error('Email send error:', err);
    res.status(500).json({ error: 'Failed to send email with PDF', details: err.message });
  }
});

const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log(`Node server running on http://localhost:${port}`);
});
