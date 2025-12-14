# Nodenberg API Test Client

GUI Test Client for Nodenberg API Server - A standalone web application for testing Excel & PDF generation APIs.

## Features

- 📁 **File Upload** - Upload Excel templates (.xlsx)
- 🔍 **Placeholder Detection** - Detect and preview placeholders in templates
- 📊 **Excel Generation** - Generate Excel files with placeholder replacement
- 📄 **PDF Generation** - Generate PDF files (requires LibreOffice on server)
- ⚙️ **API URL Configuration** - Configure API base URL via settings modal
- 💾 **LocalStorage Persistence** - API URL saved across sessions

## Prerequisites

- Node.js 18+ installed
- Nodenberg API Server running (default: http://localhost:3000)

## Installation

```bash
cd 04_api-test-client
npm install
```

## Usage

### 1. Start the Test Client

```bash
npm start
```

The client will be available at: **http://localhost:8080**

### 2. Configure API URL

Click the **⚙️ Settings** button in the top-right corner to configure the API URL:

- Default: `http://localhost:3000`
- For custom servers: Enter your API base URL (without trailing slash)
- Click **Save & Reload** to apply changes

### 3. Use the Client

1. **Upload Template** - Choose an Excel template file (.xlsx)
2. **Detect Placeholders** - Click to find placeholders like `{{company_name}}`
3. **Input Data** - Enter JSON data to replace placeholders
4. **Generate Files** - Generate Excel or PDF files
5. **Download** - Click the download link to save generated files

## Project Structure

```
04_api-test-client/
├── public/
│   ├── index.html     # Main HTML file
│   ├── style.css      # Styling
│   └── app.js         # Client logic
├── package.json       # Node.js configuration
└── README.md          # This file
```

## API URL Configuration

### Via Settings UI (Recommended)

1. Click the ⚙️ icon in the header
2. Enter the API base URL
3. Click "Save & Reload"

### Via Browser Console

```javascript
// Set custom API URL
localStorage.setItem('nodenberg_api_url', 'http://your-server:3000');
location.reload();

// Reset to default
localStorage.removeItem('nodenberg_api_url');
location.reload();
```

## Development

### Server Configuration

The test client uses `http-server` with the following settings:

- Port: 8080
- Cache: Disabled (-c-1)
- Directory: public/

### Modify Port

Edit `package.json`:

```json
{
  "scripts": {
    "start": "http-server public -p 9000 -c-1"
  }
}
```

## API Endpoints Used

The test client communicates with the following Nodenberg API endpoints:

| Endpoint | Method | Purpose |
|----------|--------|---------|
| `/health` | GET | Server health check |
| `/template/placeholders` | POST | Detect placeholders |
| `/template/info` | POST | Get template metadata |
| `/generate/excel` | POST | Generate Excel file |
| `/generate/pdf` | POST | Generate PDF file |

## CORS Configuration

The Nodenberg API Server must have CORS enabled for cross-origin requests:

```typescript
// In server.ts
app.use(cors());
```

This is already configured in the default Nodenberg API Server.

## Troubleshooting

### Connection Failed

**Error:** "Failed to connect to server"

**Solutions:**
1. Verify the API server is running: `curl http://localhost:3000/health`
2. Check the API URL in settings (⚙️ icon)
3. Ensure CORS is enabled on the API server

### PDF Generation Failed

**Error:** "LibreOffice is required for PDF generation"

**Solution:**
Install LibreOffice on the API server:

```bash
# Ubuntu/Debian
sudo apt-get install libreoffice

# macOS
brew install libreoffice

# Windows
# Download from https://www.libreoffice.org/
```

### File Upload Issues

**Error:** "Failed to read file"

**Solutions:**
1. Ensure the file is a valid Excel file (.xlsx)
2. Check file size (limit: 50MB)
3. Try a different Excel file

## License

MIT

## Related Projects

- **Nodenberg API Server** - `../03_docker-version/`
- **CLI Test Script** - `../03_docker-version/tests/test-api.js`

## Support

For issues or questions, refer to the main Nodenberg API documentation.
