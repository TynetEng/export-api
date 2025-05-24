require('dotenv').config();
const express = require('express');
const axios = require('axios');
const bodyParser = require('body-parser');
const nodemailer = require('nodemailer');

const app = express();
app.use(bodyParser.json());
const cors = require('cors');
app.use(cors());

// Get Microsoft Graph access token
const getAccessToken = async () => {
  const { CLIENT_ID, CLIENT_SECRET, TENANT_ID } = process.env;
  const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;

  const params = new URLSearchParams();
  params.append('client_id', CLIENT_ID);
  params.append('client_secret', CLIENT_SECRET);
  params.append('grant_type', 'client_credentials');
  params.append('scope', 'https://graph.microsoft.com/.default');

  // Add correct Content-Type header
  const response = await axios.post(url, params, {
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
  });
  return response.data.access_token;
};

// Get Graph site ID from host and site path
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
  // Use expand=fields to get SharePoint fields
  const response = await axios.get(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${itemId}?expand=fields`,
    {
      headers: { Authorization: `Bearer ${token}` },
    }
  );
  return response.data;
};

// Express route to fetch item by ID
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

// Debug route: List all SharePoint lists on the site
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
    // Return displayName and id for each list
    const lists = response.data.value.map(l => ({ displayName: l.displayName, id: l.id }));
    res.json(lists);
  } catch (err) {
    console.error('Error fetching SharePoint lists:', err.response?.data || err.message);
    res.status(500).json({
      error: 'Failed to fetch SharePoint lists',
      details: err.response?.data || err.message,
    });
  }
});

// Email sending endpoint
app.post('/api/send-email', async (req, res) => {
  const data = req.body;
  try {
    // Configure your SMTP transport (update with your real SMTP credentials)
    const transporter = nodemailer.createTransport({
      host: process.env.SMTP_HOST || 'smtp.example.com',
      port: process.env.SMTP_PORT ? parseInt(process.env.SMTP_PORT) : 587,
      secure: false, // true for 465, false for other ports
      auth: {
        user: process.env.SMTP_USER || 'your@email.com',
        pass: process.env.SMTP_PASS || 'yourpassword',
      },
    });

    // Compose email body (simple text version)
    let text = `Shipping Instruction Submission\n\n`;
    text += `Reference: ${data.info?.reference || ''}\n`;
    text += `Bill of Lading: ${data.info?.billOfLading || ''}\n`;
    text += `Carrier Reference: ${data.info?.carrierReference || ''}\n`;
    text += `Destination Port: ${data.info?.destinationPort || ''}\n`;
    text += `Load Point: ${data.info?.loadPoint || ''}\n`;
    text += `Vessel: ${data.info?.vessel || ''}\n`;
    text += `Local Client: ${data.info?.localClient || ''}\n`;
    text += `Pickup Date: ${data.info?.pickupDate || ''}\n\n`;
    text += `Name: ${data.name || ''}\nEmail: ${data.email || ''}\n\n`;
    text += `Consignor: ${JSON.stringify(data.consignor, null, 2)}\n\n`;
    text += `Consignee: ${JSON.stringify(data.consignee, null, 2)}\n\n`;
    text += `Notify: ${JSON.stringify(data.notify, null, 2)}\n\n`;
    text += `Shipment Value: ${data.shipmentValue || ''}\nNotes: ${data.shipmentNotes || ''}\n\n`;
    text += `Containers:\n`;
    (data.containers || []).forEach((c, i) => {
      text += `  [${i + 1}] ${JSON.stringify(c, null, 2)}\n`;
    });

    await transporter.sendMail({
      from: process.env.SMTP_FROM || 'noreply@yourdomain.com',
      to: process.env.SMTP_TO || 'your@email.com',
      subject: 'New Shipping Instruction Submission',
      text,
    });
    res.json({ success: true });
  } catch (err) {
    console.error('Email send error:', err);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.listen(3000, () => {
  console.log('Node server running on http://localhost:3000');
});
