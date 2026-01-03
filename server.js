// ============================================================================
// EXCHANGE ONLINE SHARED MAILBOX API - For Render Deployment
// Creates TRUE shared mailboxes using PowerShell Core + Exchange Online Module
// ============================================================================

const express = require('express');
const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');
const os = require('os');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(express.json());

// API Key authentication
const API_KEY = process.env.API_KEY || 'atomicwork-secret-key-change-this';

// Azure AD App credentials (from environment variables)
const CONFIG = {
  clientId: process.env.AZURE_CLIENT_ID || 'your-azure-client-id',
  tenantId: process.env.AZURE_TENANT_ID || 'your-azure-tenant-id',
  certBase64: process.env.AZURE_CERT_BASE64 || '',
  certPassword: process.env.AZURE_CERT_PASSWORD || '',
  organization: process.env.EXCHANGE_ORG || 'your-organization.onmicrosoft.com'
};

// Authentication middleware
function authenticate(req, res, next) {
  const apiKey = req.headers['x-api-key'] || req.headers['authorization']?.replace('Bearer ', '');
  
  if (!apiKey || apiKey !== API_KEY) {
    return res.status(401).json({
      success: false,
      error: 'Unauthorized. Provide valid X-API-Key header.'
    });
  }
  next();
}

// ============================================================================
// PowerShell Execution Helper
// ============================================================================
async function runPowerShell(script) {
  return new Promise((resolve, reject) => {
    const tempFile = path.join(os.tmpdir(), `exchange_${Date.now()}.ps1`);
    
    // Wrap script with connection logic
    const fullScript = `
$ErrorActionPreference = "Continue"
$result = @{
  success = $false
  data = $null
  error = $null
  logs = @()
}
$tempCertPath = $null

try {
    $result.logs += "Importing Exchange Online module..."
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    
    $result.logs += "Connecting to Exchange Online..."
    
    # Connect using Certificate
    $certBytes = [Convert]::FromBase64String("${CONFIG.certBase64}")
    $tempCertPath = "/tmp/exchange_cert.pfx"
    [System.IO.File]::WriteAllBytes($tempCertPath, $certBytes)
    
    Connect-ExchangeOnline -AppId "${CONFIG.clientId}" -Organization "${CONFIG.organization}" -CertificateFilePath $tempCertPath -CertificatePassword (ConvertTo-SecureString "${CONFIG.certPassword}" -AsPlainText -Force) -ShowBanner:$false -ErrorAction Stop
    
    $result.logs += "Connected successfully!"
    
    # Execute the actual script
    ${script}
    
    $result.success = $true
}
catch {
    $result.error = $_.Exception.Message
    $result.logs += "ERROR: $($_.Exception.Message)"
}
finally {
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        if ($tempCertPath -and (Test-Path $tempCertPath)) { Remove-Item $tempCertPath -Force }
        $result.logs += "Disconnected from Exchange Online"
    } catch {}
}

# Output as JSON
$result | ConvertTo-Json -Depth 10 -Compress
`;

    fs.writeFileSync(tempFile, fullScript, 'utf8');
    
    const pwsh = spawn('pwsh', ['-NoProfile', '-NonInteractive', '-File', tempFile], {
      timeout: 120000
    });
    
    let stdout = '';
    let stderr = '';
    
    pwsh.stdout.on('data', (data) => {
      stdout += data.toString();
    });
    
    pwsh.stderr.on('data', (data) => {
      stderr += data.toString();
    });
    
    pwsh.on('close', (code) => {
      // Clean up temp file
      try { fs.unlinkSync(tempFile); } catch (e) {}
      
      // Try to parse JSON from stdout
      try {
        // Find JSON in output
        const jsonMatch = stdout.match(/\{.*\}/s);
        if (jsonMatch) {
          const result = JSON.parse(jsonMatch[0]);
          resolve(result);
        } else {
          resolve({
            success: code === 0,
            data: null,
            error: stderr || 'No JSON output',
            logs: [stdout, stderr].filter(Boolean),
            rawOutput: stdout
          });
        }
      } catch (parseErr) {
        resolve({
          success: false,
          error: `Parse error: ${parseErr.message}`,
          logs: [stdout, stderr].filter(Boolean),
          rawOutput: stdout
        });
      }
    });
    
    pwsh.on('error', (err) => {
      try { fs.unlinkSync(tempFile); } catch (e) {}
      reject(err);
    });
  });
}


// ============================================================================
// ENDPOINT 1: Create Shared Mailbox
// POST /api/shared-mailbox/create
// ============================================================================
app.post('/api/shared-mailbox/create', authenticate, async (req, res) => {
  console.log('=== Create Shared Mailbox ===');
  console.log('Body:', JSON.stringify(req.body, null, 2));
  
  const {
    mailboxName,
    mailboxEmail,
    mailboxAlias,
    members = '',
    fullAccess = true,
    sendAs = true,
    sendOnBehalf = true,
    autoMapping = true
  } = req.body;
  
  // Validation
  if (!mailboxName) {
    return res.status(400).json({ success: false, error: 'mailboxName is required' });
  }
  if (!mailboxEmail) {
    return res.status(400).json({ success: false, error: 'mailboxEmail is required' });
  }
  
  // Parse members
  const memberList = typeof members === 'string' 
    ? members.split(',').map(m => m.trim()).filter(Boolean)
    : (Array.isArray(members) ? members : []);
  
  // Build PowerShell script
  let psScript = `
    $result.logs += "Creating shared mailbox: ${mailboxEmail}"
    
    # Check if mailbox already exists
    $existing = Get-Mailbox -Identity "${mailboxEmail}" -ErrorAction SilentlyContinue
    if ($existing) {
        throw "Mailbox '${mailboxEmail}' already exists"
    }
    
    # Create the shared mailbox
    $mailbox = New-Mailbox -Shared -Name "${mailboxName}" -DisplayName "${mailboxName}" -PrimarySmtpAddress "${mailboxEmail}" -ErrorAction Stop
    
    $result.logs += "Mailbox created successfully!"
    $result.data = @{
        id = $mailbox.ExternalDirectoryObjectId
        name = $mailbox.Name
        displayName = $mailbox.DisplayName
        email = $mailbox.PrimarySmtpAddress
        alias = $mailbox.Alias
        type = "SharedMailbox"
        membersAdded = @()
        permissionsGranted = @()
    }
  `;
  
  // Add permissions for each member
  for (const member of memberList) {
    if (fullAccess) {
      psScript += `
    # Grant FullAccess to ${member}
    try {
        Add-MailboxPermission -Identity "${mailboxEmail}" -User "${member}" -AccessRights FullAccess -InheritanceType All ${autoMapping ? '' : '-AutoMapping $false'} -ErrorAction Stop | Out-Null
        $result.data.permissionsGranted += "FullAccess:${member}"
        $result.logs += "FullAccess granted to ${member}"
    } catch {
        $result.logs += "Failed FullAccess for ${member}: $($_.Exception.Message)"
    }
      `;
    }
    
    if (sendAs) {
      psScript += `
    # Grant SendAs to ${member}
    try {
        Add-RecipientPermission -Identity "${mailboxEmail}" -Trustee "${member}" -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
        $result.data.permissionsGranted += "SendAs:${member}"
        $result.logs += "SendAs granted to ${member}"
    } catch {
        $result.logs += "Failed SendAs for ${member}: $($_.Exception.Message)"
    }
      `;
    }
    
    if (sendOnBehalf) {
      psScript += `
    # Grant SendOnBehalf to ${member}
    try {
        Set-Mailbox -Identity "${mailboxEmail}" -GrantSendOnBehalfTo @{Add="${member}"} -ErrorAction Stop
        $result.data.permissionsGranted += "SendOnBehalf:${member}"
        $result.logs += "SendOnBehalf granted to ${member}"
    } catch {
        $result.logs += "Failed SendOnBehalf for ${member}: $($_.Exception.Message)"
    }
      `;
    }
    
    psScript += `$result.data.membersAdded += "${member}"
    `;
  }
  
  try {
    const result = await runPowerShell(psScript);
    
    const response = {
      success: result.success,
      sharedMailbox: result.data || {
        name: mailboxName,
        email: mailboxEmail
      },
      members: memberList,
      permissions: { fullAccess, sendAs, sendOnBehalf, autoMapping },
      logs: result.logs,
      error: result.error,
      agentMessage: result.success 
        ? `✅ Shared mailbox "${mailboxName}" (${mailboxEmail}) created successfully with ${memberList.length} member(s).` 
        : `❌ Failed to create shared mailbox: ${result.error}` 
    };
    
    res.status(result.success ? 200 : 500).json(response);
  } catch (err) {
    console.error('Error:', err);
    res.status(500).json({
      success: false,
      error: err.message,
      agentMessage: `❌ Failed to create shared mailbox: ${err.message}` 
    });
  }
});


// ============================================================================
// ENDPOINT 2: Add Members to Shared Mailbox
// POST /api/shared-mailbox/add-members
// ============================================================================
app.post('/api/shared-mailbox/add-members', authenticate, async (req, res) => {
  console.log('=== Add Members ===');
  console.log('Body:', JSON.stringify(req.body, null, 2));
  
  const {
    mailboxEmail,
    members = '',
    fullAccess = true,
    sendAs = true,
    sendOnBehalf = true,
    autoMapping = true
  } = req.body;
  
  if (!mailboxEmail) {
    return res.status(400).json({ success: false, error: 'mailboxEmail is required' });
  }
  
  const memberList = typeof members === 'string' 
    ? members.split(',').map(m => m.trim()).filter(Boolean)
    : (Array.isArray(members) ? members : []);
  
  if (memberList.length === 0) {
    return res.status(400).json({ success: false, error: 'At least one member is required' });
  }
  
  let psScript = `
    $result.logs += "Adding members to: ${mailboxEmail}"
    
    # Verify mailbox exists
    $mailbox = Get-Mailbox -Identity "${mailboxEmail}" -ErrorAction Stop
    $result.logs += "Found mailbox: $($mailbox.DisplayName)"
    
    $result.data = @{
        mailbox = $mailbox.PrimarySmtpAddress
        displayName = $mailbox.DisplayName
        membersAdded = @()
        permissionsGranted = @()
        errors = @()
    }
  `;
  
  for (const member of memberList) {
    if (fullAccess) {
      psScript += `
    try {
        Add-MailboxPermission -Identity "${mailboxEmail}" -User "${member}" -AccessRights FullAccess -InheritanceType All ${autoMapping ? '' : '-AutoMapping $false'} -ErrorAction Stop | Out-Null
        $result.data.permissionsGranted += "FullAccess:${member}"
        $result.logs += "FullAccess granted to ${member}"
    } catch {
        $result.data.errors += "FullAccess:${member}:$($_.Exception.Message)"
    }
      `;
    }
    
    if (sendAs) {
      psScript += `
    try {
        Add-RecipientPermission -Identity "${mailboxEmail}" -Trustee "${member}" -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
        $result.data.permissionsGranted += "SendAs:${member}"
        $result.logs += "SendAs granted to ${member}"
    } catch {
        $result.data.errors += "SendAs:${member}:$($_.Exception.Message)"
    }
      `;
    }
    
    if (sendOnBehalf) {
      psScript += `
    try {
        Set-Mailbox -Identity "${mailboxEmail}" -GrantSendOnBehalfTo @{Add="${member}"} -ErrorAction Stop
        $result.data.permissionsGranted += "SendOnBehalf:${member}"
        $result.logs += "SendOnBehalf granted to ${member}"
    } catch {
        $result.data.errors += "SendOnBehalf:${member}:$($_.Exception.Message)"
    }
      `;
    }
    
    psScript += `$result.data.membersAdded += "${member}"
    `;
  }
  
  try {
    const result = await runPowerShell(psScript);
    
    res.status(result.success ? 200 : 500).json({
      success: result.success,
      mailbox: mailboxEmail,
      membersAdded: memberList,
      data: result.data,
      logs: result.logs,
      error: result.error,
      agentMessage: result.success 
        ? `✅ Added ${memberList.length} member(s) to "${mailboxEmail}".` 
        : `❌ Failed to add members: ${result.error}` 
    });
  } catch (err) {
    res.status(500).json({
      success: false,
      error: err.message,
      agentMessage: `❌ Failed to add members: ${err.message}` 
    });
  }
});


// ============================================================================
// ENDPOINT 3: Remove Members from Shared Mailbox
// POST /api/shared-mailbox/remove-members
// ============================================================================
app.post('/api/shared-mailbox/remove-members', authenticate, async (req, res) => {
  console.log('=== Remove Members ===');
  console.log('Body:', JSON.stringify(req.body, null, 2));
  
  const {
    mailboxEmail,
    members = '',
    removeFullAccess = true,
    removeSendAs = true,
    removeSendOnBehalf = true,
    reason = 'Access revocation'
  } = req.body;
  
  if (!mailboxEmail) {
    return res.status(400).json({ success: false, error: 'mailboxEmail is required' });
  }
  
  const memberList = typeof members === 'string' 
    ? members.split(',').map(m => m.trim()).filter(Boolean)
    : (Array.isArray(members) ? members : []);
  
  if (memberList.length === 0) {
    return res.status(400).json({ success: false, error: 'At least one member is required' });
  }
  
  let psScript = `
    $result.logs += "Removing members from: ${mailboxEmail}"
    $result.logs += "Reason: ${reason}"
    
    # Verify mailbox exists
    $mailbox = Get-Mailbox -Identity "${mailboxEmail}" -ErrorAction Stop
    $result.logs += "Found mailbox: $($mailbox.DisplayName)"
    
    $result.data = @{
        mailbox = $mailbox.PrimarySmtpAddress
        displayName = $mailbox.DisplayName
        membersRemoved = @()
        permissionsRevoked = @()
        errors = @()
    }
  `;
  
  for (const member of memberList) {
    if (removeFullAccess) {
      psScript += `
    try {
        Remove-MailboxPermission -Identity "${mailboxEmail}" -User "${member}" -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction Stop | Out-Null
        $result.data.permissionsRevoked += "FullAccess:${member}"
        $result.logs += "FullAccess removed from ${member}"
    } catch {
        $result.data.errors += "FullAccess:${member}:$($_.Exception.Message)"
    }
      `;
    }
    
    if (removeSendAs) {
      psScript += `
    try {
        Remove-RecipientPermission -Identity "${mailboxEmail}" -Trustee "${member}" -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
        $result.data.permissionsRevoked += "SendAs:${member}"
        $result.logs += "SendAs removed from ${member}"
    } catch {
        $result.data.errors += "SendAs:${member}:$($_.Exception.Message)"
    }
      `;
    }
    
    if (removeSendOnBehalf) {
      psScript += `
    try {
        Set-Mailbox -Identity "${mailboxEmail}" -GrantSendOnBehalfTo @{Remove="${member}"} -ErrorAction Stop
        $result.data.permissionsRevoked += "SendOnBehalf:${member}"
        $result.logs += "SendOnBehalf removed from ${member}"
    } catch {
        $result.data.errors += "SendOnBehalf:${member}:$($_.Exception.Message)"
    }
      `;
    }
    
    psScript += `$result.data.membersRemoved += "${member}"
    `;
  }
  
  try {
    const result = await runPowerShell(psScript);
    
    res.status(result.success ? 200 : 500).json({
      success: result.success,
      mailbox: mailboxEmail,
      membersRemoved: memberList,
      reason: reason,
      data: result.data,
      logs: result.logs,
      error: result.error,
      agentMessage: result.success 
        ? `🔒 Removed ${memberList.length} member(s) from "${mailboxEmail}". Reason: ${reason}` 
        : `❌ Failed to remove members: ${result.error}` 
    });
  } catch (err) {
    res.status(500).json({
      success: false,
      error: err.message,
      agentMessage: `❌ Failed to remove members: ${err.message}` 
    });
  }
});


// ============================================================================
// ENDPOINT 4: List Mailbox Permissions
// GET /api/shared-mailbox/permissions?email=support@domain.com
// ============================================================================
app.get('/api/shared-mailbox/permissions', authenticate, async (req, res) => {
  const mailboxEmail = req.query.email;
  
  if (!mailboxEmail) {
    return res.status(400).json({ success: false, error: 'email query parameter is required' });
  }
  
  const psScript = `
    $result.logs += "Getting permissions for: ${mailboxEmail}"
    
    $mailbox = Get-Mailbox -Identity "${mailboxEmail}" -ErrorAction Stop
    
    $result.data = @{
        mailbox = @{
            displayName = $mailbox.DisplayName
            email = $mailbox.PrimarySmtpAddress
            type = $mailbox.RecipientTypeDetails
        }
        fullAccess = @()
        sendAs = @()
        sendOnBehalf = @()
    }
    
    # Get FullAccess permissions
    $fullAccess = Get-MailboxPermission -Identity "${mailboxEmail}" | Where-Object { $_.User -notlike "NT AUTHORITY*" -and $_.User -notlike "S-1-5*" -and $_.AccessRights -contains "FullAccess" }
    $result.data.fullAccess = @($fullAccess | ForEach-Object { $_.User.ToString() })
    
    # Get SendAs permissions  
    $sendAs = Get-RecipientPermission -Identity "${mailboxEmail}" | Where-Object { $_.Trustee -notlike "NT AUTHORITY*" }
    $result.data.sendAs = @($sendAs | ForEach-Object { $_.Trustee.ToString() })
    
    # Get SendOnBehalf permissions
    $result.data.sendOnBehalf = @($mailbox.GrantSendOnBehalfTo | ForEach-Object { $_.ToString() })
    
    $result.logs += "Permissions retrieved"
  `;
  
  try {
    const result = await runPowerShell(psScript);
    res.json({
      success: result.success,
      ...result.data,
      logs: result.logs,
      error: result.error
    });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});


// ============================================================================
// ENDPOINT 5: Delete Shared Mailbox
// DELETE /api/shared-mailbox?email=support@domain.com
// ============================================================================
app.delete('/api/shared-mailbox', authenticate, async (req, res) => {
  const mailboxEmail = req.query.email || req.body.mailboxEmail;
  const confirm = req.body.confirm === true;
  
  if (!mailboxEmail) {
    return res.status(400).json({ success: false, error: 'email is required' });
  }
  
  if (!confirm) {
    return res.status(400).json({ 
      success: false, 
      error: 'Please confirm deletion by setting confirm: true in request body' 
    });
  }
  
  const psScript = `
    $result.logs += "Deleting shared mailbox: ${mailboxEmail}"
    
    $mailbox = Get-Mailbox -Identity "${mailboxEmail}" -ErrorAction Stop
    $result.data = @{
        deletedMailbox = $mailbox.DisplayName
        email = $mailbox.PrimarySmtpAddress
    }
    
    Remove-Mailbox -Identity "${mailboxEmail}" -Confirm:$false -ErrorAction Stop
    $result.logs += "Mailbox deleted successfully"
  `;
  
  try {
    const result = await runPowerShell(psScript);
    res.json({
      success: result.success,
      deleted: mailboxEmail,
      data: result.data,
      logs: result.logs,
      error: result.error,
      agentMessage: result.success 
        ? `🗑️ Shared mailbox "${mailboxEmail}" deleted.` 
        : `❌ Failed to delete: ${result.error}` 
    });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});


// ============================================================================
// Health Check
// ============================================================================
app.get('/health', (req, res) => {
  res.json({ 
    status: 'healthy', 
    service: 'Exchange Online Shared Mailbox API',
    timestamp: new Date().toISOString(),
    environment: process.env.NODE_ENV || 'development'
  });
});

// Root endpoint
app.get('/', (req, res) => {
  res.json({
    service: 'Exchange Online Shared Mailbox API',
    version: '1.0.0',
    endpoints: [
      'POST /api/shared-mailbox/create',
      'POST /api/shared-mailbox/add-members',
      'POST /api/shared-mailbox/remove-members',
      'GET  /api/shared-mailbox/permissions?email=...',
      'DELETE /api/shared-mailbox?email=...',
      'GET  /health'
    ],
    documentation: 'Include X-API-Key header for authentication'
  });
});


// ============================================================================
// Start Server
// ============================================================================
app.listen(PORT, () => {
  console.log(`
╔═══════════════════════════════════════════════════════════════╗
║     Exchange Online Shared Mailbox API - Running on Render    ║
╠═══════════════════════════════════════════════════════════════╣
║  Port: ${PORT}                                                    ║
║  Environment: ${(process.env.NODE_ENV || 'development').padEnd(44)}║
║                                                               ║
║  Endpoints:                                                   ║
║  POST   /api/shared-mailbox/create        Create mailbox      ║
║  POST   /api/shared-mailbox/add-members   Add members         ║
║  POST   /api/shared-mailbox/remove-members Remove members     ║
║  GET    /api/shared-mailbox/permissions   List permissions    ║
║  DELETE /api/shared-mailbox               Delete mailbox      ║
╚═══════════════════════════════════════════════════════════════╝
  `);
});

module.exports = app;
