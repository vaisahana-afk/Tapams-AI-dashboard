# TAPAMS Turnaround Intelligence

Deployment handoff for the TAPAMS customer and product intelligence prototype.

## Current App

The application is a static HTML/CSS/JavaScript app with these pages:

- `index.html` - TAPAMS sign-in page
- `upload.html` - Upload Excel, Word, CSV, PDF, or text files
- `dashboard.html` - AI Intelligence dashboard
- `tapams-logo.svg` - TAPAMS Group of Companies logo
- `styles.css` - Shared styling
- `script.js` - Sign-in behavior
- `upload.js` - File upload, sample Excel download, and routing
- `dashboard.js` - File reading, category analysis, effective-date analysis, and customer risk intelligence

## Run Locally

From the project folder:

```powershell
python -m http.server 8000
```

Open:

```text
http://127.0.0.1:8000/index.html
```

Upload page:

```text
http://127.0.0.1:8000/upload.html
```

Dashboard page:

```text
http://127.0.0.1:8000/dashboard.html
```

## Deploy to Netlify

### Option 1: Netlify web interface

1. Open https://app.netlify.com/drop
2. Sign in or create a Netlify account.
3. Drag the entire `D:\Sample-Python` folder into the upload area.
4. Netlify will generate a public URL.
5. Rename the site in Netlify if desired, for example `tapams-turnaround-intelligence`.
6. Share the generated URL, such as:

```text
https://tapams-turnaround-intelligence.netlify.app
```

### Option 2: Netlify CLI

Install and sign in:

```powershell
npm install -g netlify-cli
netlify login
```

From `D:\Sample-Python`, create a site:

```powershell
netlify init
```

Deploy a draft:

```powershell
netlify deploy
```

Deploy to production:

```powershell
netlify deploy --prod
```

The CLI will show the public URL after deployment.

## Important Browser Storage Note

Uploaded files are currently stored in the browser's IndexedDB and analyzed locally in that browser. This means:

- The uploaded file is not sent to a server.
- Another person opening the public URL will see the upload page, not your uploaded file.
- Each user must upload their own Excel or Word file.
- The current prototype does not have shared accounts, server storage, or server-side AI processing.

## Sample Excel

On the upload page, click **Download sample effective-dated Excel**.

The browser downloads:

```text
tapams-effective-dated-sample.xlsx
```

The sample includes:

- Customer Code
- Customer Name
- Value
- Category
- Product
- Invoice Date
- Effective From
- Effective To
- Last Enhancement Date

## Intelligence Currently Supported

The dashboard provides:

- Total value
- Customer account count
- CCCP and NCNP category analysis
- Customer category selector
- Category-specific value and customer details
- Invoice-date and effective-date comparison
- Active and expired product status
- Product enhancement aging
- Portfolio concentration risk
- Customer-level High, Watch, Opportunity, and Low ratings
- Actionable revenue recommendations for NCNP accounts
- Renewal and replacement recommendations for expired CCCP products

## Category Definitions

- `CCCP` - Current Customer / Current Product
- `NCNP` - New Customer / New Product

## Risk Definitions

### High
Urgent revenue risk, such as an expired product validity period, uncertain renewal, or significant revenue exposure.

### Watch
The customer or product needs monitoring, often because of CCCP dependency, enhancement aging, or renewal exposure.

### Opportunity
The account has growth potential, commonly from NCNP activity. Recommended actions include repeat orders, cross-selling, customer reviews, and account expansion.

### Low
No immediate risk signal was detected from the available data.

## Recommended Next Steps

1. Deploy the static prototype to Netlify.
2. Upload the real Excel workbook and Word requirement document.
3. Validate the category and customer intelligence with business stakeholders.
4. Add a backend and authentication before using real confidential customer data.
5. Add server-side document parsing and an approved AI service for production intelligence.
