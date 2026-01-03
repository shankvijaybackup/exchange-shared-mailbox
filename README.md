# Exchange Online Shared Mailbox API

REST API for managing **true shared mailboxes** in Microsoft 365/Exchange Online. Designed for Atomicwork integration via webhooks.

## 🚀 Quick Deploy to Render

### Step 1: Push to GitHub

```bash
# Create a new repo or use existing
git init
git add .
git commit -m "Exchange Mailbox API"
git remote add origin https://github.com/YOUR_USERNAME/exchange-mailbox-api.git
git push -u origin main
```

### Step 2: Deploy on Render

1. Go to [Render Dashboard](https://dashboard.render.com)
2. Click **New** → **Web Service**
3. Connect your GitHub repository
4. Render will auto-detect the Dockerfile
5. Configure:
   - **Name:** `exchange-mailbox-api` 
   - **Region:** Choose closest to you
   - **Plan:** Starter ($7/month) or Standard ($25/month)

### Step 3: Set Environment Variables

In Render Dashboard → Your Service → **Environment**:

| Variable | Value |
|----------|-------|
| `API_KEY` | `your-secure-api-key-here` (generate a strong random string) |
| `AZURE_CLIENT_ID` | `your-azure-client-id` |
| `AZURE_TENANT_ID` | `your-azure-tenant-id` |
| `AZURE_CLIENT_SECRET` | `your-azure-client-secret` |
| `EXCHANGE_ORG` | `your-organization.onmicrosoft.com` |

### Step 4: Azure AD App Permissions

Your Azure AD app needs these **Application permissions**:

1. Go to Azure Portal → Azure AD → App Registrations → Your App
2. API Permissions → Add Permission → APIs my organization uses
3. Search for **Office 365 Exchange Online**
4. Add: `Exchange.ManageAsApp` 
5. **Grant admin consent**

Also assign **Exchange Administrator** role:
1. Azure AD → Roles and administrators
2. Find "Exchange Administrator"
3. Add your app's Service Principal

---

## 📡 API Endpoints

Base URL: `https://your-app.onrender.com` 

### Authentication
All requests require `X-API-Key` header:
```
X-API-Key: your-secure-api-key
```

---

### 1. Create Shared Mailbox

**POST** `/api/shared-mailbox/create` 

```json
{
  "mailboxName": "Support Team",
  "mailboxEmail": "support@atombank.co",
  "members": "vijay@atombank.co, john@atombank.co",
  "fullAccess": true,
  "sendAs": true,
  "sendOnBehalf": true,
  "autoMapping": true
}
```

**Response:**
```json
{
  "success": true,
  "sharedMailbox": {
    "name": "Support Team",
    "email": "support@atombank.co",
    "type": "SharedMailbox"
  },
  "members": ["vijay@atombank.co", "john@atombank.co"],
  "agentMessage": "✅ Shared mailbox created successfully..."
}
```

---

### 2. Add Members

**POST** `/api/shared-mailbox/add-members` 

```json
{
  "mailboxEmail": "support@atombank.co",
  "members": "newuser@atombank.co",
  "fullAccess": true,
  "sendAs": true,
  "sendOnBehalf": true
}
```

---

### 3. Remove Members

**POST** `/api/shared-mailbox/remove-members` 

```json
{
  "mailboxEmail": "support@atombank.co",
  "members": "olduser@atombank.co",
  "reason": "Employee offboarding"
}
```

---

### 4. Get Permissions

**GET** `/api/shared-mailbox/permissions?email=support@atombank.co` 

---

### 5. Delete Mailbox

**DELETE** `/api/shared-mailbox?email=support@atombank.co` 

Body:
```json
{
  "confirm": true
}
```

---

## 🔗 Atomicwork Integration

### Configure Webhook Action

1. In Atomicwork, create a new **HTTP Request** action
2. Configure:
   - **Method:** POST
   - **URL:** `https://your-app.onrender.com/api/shared-mailbox/create` 
   - **Headers:**
     ```
     Content-Type: application/json
     X-API-Key: {{secrets.EXCHANGE_API_KEY}}
     ```
   - **Body:**
     ```json
     {
       "mailboxName": "{{input.mailboxName}}",
       "mailboxEmail": "{{input.mailboxEmail}}",
       "members": "{{input.members}}",
       "fullAccess": true,
       "sendAs": true,
       "sendOnBehalf": true
     }
     ```

---

## 🔒 Security Notes

1. **Rotate API keys** regularly
2. **Use HTTPS only** (Render provides this automatically)
3. **Limit IP access** if possible (Render paid plans)
4. **Monitor logs** for suspicious activity
5. **Keep Azure credentials secure** - never commit to git

---

## 🐛 Troubleshooting

### "Connection failed" error
- Check Azure AD app has `Exchange.ManageAsApp` permission
- Verify Exchange Administrator role is assigned
- Ensure admin consent is granted

### "Mailbox not found" error
- Verify email address format
- Check mailbox exists in Exchange Admin Center

### PowerShell errors
- Check Render logs: Dashboard → Your Service → Logs
- Verify ExchangeOnlineManagement module installed correctly

---

## 📊 Monitoring

Render provides:
- Real-time logs
- Health checks every 30 seconds
- Auto-restart on failure
- Metrics dashboard

---

## 💰 Render Pricing

| Plan | Price | Best For |
|------|-------|----------|
| Starter | $7/month | Testing, low volume |
| Standard | $25/month | Production |
| Pro | $85/month | High availability |

---

## Support

- Render Docs: https://render.com/docs
- Exchange Online PS: https://docs.microsoft.com/powershell/exchange/
- Graph API: https://docs.microsoft.com/graph/
