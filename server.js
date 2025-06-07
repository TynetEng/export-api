require('dotenv').config();
const express = require('express');
const cors = require('cors');
const axios = require('axios');
const bodyParser = require('body-parser');

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

app.listen(3000, () => {
  console.log('Node server running on http://localhost:3000');
});

// Export app for testing
