

const express = require('express');
const path = require('path');
const fetch = require('node-fetch');
require('dotenv').config();

const app = express();
app.use(express.json());

const PORT = process.env.PORT || 3000;
const GAS_URL = process.env.GAS_URL;

if (!GAS_URL) {
  console.error("FATAL ERROR: GAS_URL environment variable is not set.");
  process.exit(1);
}

async function processImageUrlsToBase64(data) {
  if (!data || typeof data !== 'object') {
    return;
  }
  
  if (Array.isArray(data)) {
    for (const item of data) {
      await processImageUrlsToBase64(item);
    }
  } else {
    for (const key in data) {
      if (Object.prototype.hasOwnProperty.call(data, key)) {
        if (key === 'PhotoURL' && typeof data[key] === 'string' && (data[key].startsWith('http://') || data[key].startsWith('https://'))) {
          try {
            const response = await fetch(data[key]);
            if (response.ok) {
              const imageBuffer = await response.buffer();
              const contentType = response.headers.get('content-type') || 'image/jpeg';
              data[key] = `data:${contentType};base64,${imageBuffer.toString('base64')}`;
            } else {
              console.warn(`Failed to fetch image ${data[key]}: Status ${response.status}`);
            }
          } catch (error) {
            console.error(`Error converting image URL ${data[key]} to Base64:`, error.message);
          }
        }
        
        if (typeof data[key] === 'object') {
          await processImageUrlsToBase64(data[key]);
        }
      }
    }
  }
}

app.get('/', (req, res) => {
  const indexPath = path.join(__dirname, 'index.html');
  res.sendFile(indexPath, (err) => {
    if (err) {
      console.error("Error sending index.html:", err);
      if (!res.headersSent) {
        res.status(500).send("Error loading application shell.");
      }
    }
  });
});

app.get('/tools.json', (req, res) => {
  const toolsPath = path.join(__dirname, 'tools.json');
  res.sendFile(toolsPath, (err) => {
    if (err) {
      console.error("Error sending tools.json:", err);
      if (!res.headersSent) {
        res.status(err.status || 500).send("Error loading tools configuration.");
      }
    }
  });
});

app.get('/generator.json', (req, res) => {
  const generatorPath = path.join(__dirname, 'generator.json');
  res.sendFile(generatorPath, (err) => {
    if (err) {
      console.error("Error sending generator.json:", err);
      if (!res.headersSent) {
        res.status(err.status || 500).send("Error loading generator configuration.");
      }
    }
  });
});


app.post('/login', async (req, res) => {
  const { mobile, password } = req.body;
  if (!mobile || !password) {
    return res.status(400).json({ success: false, error: 'Mobile and password required' });
  }
  
  try {
    const gasResponse = await fetch(GAS_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ action: 'login', mobile, password })
    });
    
    const responseBodyText = await gasResponse.text();
    let result;
    try {
      result = JSON.parse(responseBodyText);
    } catch (parseError) {
      console.error('Failed to parse GAS response for login:', responseBodyText.substring(0, 500));
      return res.status(500).json({ success: false, error: `GAS returned non-JSON response. Status: ${gasResponse.status}` });
    }
    
    if (!gasResponse.ok) {
      console.error(`GAS login request failed with HTTP Status: ${gasResponse.status}`, result);
      return res.status(gasResponse.status || 502).json(result || { success: false, error: `Upstream GAS error: ${gasResponse.status}` });
    }
    
    res.status(200).json(result);
    
  } catch (error) {
    console.error('Login proxy error:', error);
    res.status(500).json({ success: false, error: 'Proxy server error during login.', details: error.message });
  }
});

app.post('/api/:action', async (req, res) => {
  const action = req.params.action;
  const payload = req.body || {};
  
  if (!action) {
    return res.status(400).json({ success: false, error: 'Action parameter is required' });
  }
  
  try {
    const bodyToGAS = JSON.stringify({ action, ...payload });
    
    const gasResponse = await fetch(GAS_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: bodyToGAS
    });
    
    const responseBodyText = await gasResponse.text();
    let result;
    try {
      result = JSON.parse(responseBodyText);
    } catch (parseError) {
      console.error(`Failed to parse GAS response for action "${action}":`, responseBodyText.substring(0, 500));
      return res.status(500).json({ success: false, error: `GAS returned non-JSON response for action ${action}. Status: ${gasResponse.status}` });
    }
    
    if (!gasResponse.ok) {
      console.error(`GAS API request failed for action "${action}" with HTTP Status: ${gasResponse.status}`, result);
      return res.status(gasResponse.status || 502).json(result || { success: false, error: `Upstream GAS error for ${action}: ${gasResponse.status}` });
    }
    
    if (gasResponse.ok && result.success && result.data) {
      await processImageUrlsToBase64(result.data);
    }
    
    res.status(200).json(result);
    
  } catch (error) {
    console.error(`API proxy error for action "${action}":`, error);
    res.status(500).json({ success: false, error: `Proxy server error during action: ${action}.`, details: error.message });
  }
});

app.get('/health', (req, res) => {
  res.status(200).send('OK');
});

app.use((req, res) => {
  if (!res.headersSent) {
    res.status(404).send("Resource not found on proxy.");
  }
});

app.listen(PORT, () => {
  console.log(`Node.js proxy server listening on port ${PORT}`);
  console.log(`Forwarding requests to GAS URL: ${GAS_URL}`);
});