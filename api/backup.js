export default async function handler(req, res) {
  // Set CORS headers to allow your study website to connect
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  
  // Handle preflight requests
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }
  
  // Only accept POST requests
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    // Get the data from your study
    const { action, data, sessionId, fileHint } = req.body;
    
    // Your GitHub details
    const GITHUB_TOKEN = process.env.GITHUB_TOKEN;
    const owner = 'melodyfschwenk';
    const repo = 'readingcomp';
    
    // Clean up the file path
    const safePath = (fileHint || `data/sessions/backup_${Date.now()}.json`)
      .replace(/\.\./g, '')
      .replace(/\/+/g, '/');
    
    // Convert data to base64 (GitHub requires this format)
    const content = Buffer.from(JSON.stringify(data, null, 2)).toString('base64');
    
    // Try to get existing file first (to get its SHA if it exists)
    let sha;
    try {
      const getResponse = await fetch(
        `https://api.github.com/repos/${owner}/${repo}/contents/${safePath}`,
        {
          headers: {
            'Authorization': `Bearer ${GITHUB_TOKEN}`,
          }
        }
      );
      if (getResponse.ok) {
        const existingFile = await getResponse.json();
        sha = existingFile.sha;
      }
    } catch (e) {
      // File doesn't exist yet, that's okay
    }
    
    // Save to GitHub
    const response = await fetch(
      `https://api.github.com/repos/${owner}/${repo}/contents/${safePath}`,
      {
        method: 'PUT',
        headers: {
          'Authorization': `Bearer ${GITHUB_TOKEN}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          message: `Add data: ${action} for session ${sessionId}`,
          content: content,
          sha: sha // Include SHA if file exists (for updates)
        })
      }
    );
    
    if (response.ok) {
      console.log('Successfully saved to GitHub:', safePath);
      res.status(200).json({ success: true, path: safePath });
    } else {
      const error = await response.text();
      console.error('GitHub API error:', error);
      res.status(500).json({ error: 'Failed to save to GitHub', details: error });
    }
  } catch (error) {
    console.error('Server error:', error);
    res.status(500).json({ error: error.message });
  }
}
