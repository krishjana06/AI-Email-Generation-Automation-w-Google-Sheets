//    Library Name: OAuth2
//    Key: *** (YOUR OAUTH2 KEY)

// Runs when the spreadsheet is opened; adds the “Auto-Email” menu
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Auto-Email')
    .addItem('Draft Emails','showSidebar')
    .addToUi();
}

// Shows the sidebar UI
function showSidebar() {
  var html = HtmlService
    .createHtmlOutputFromFile('Sidebar')
    .setWidth(300)
    .setHeight(400)
    .setTitle('Auto-Email Generator');
  SpreadsheetApp.getUi().showSidebar(html);
}

// Main entrypoint called from sidebar
// templateVariant and toneLevel are optional parameters
function generateDrafts(templateVariant, toneLevel) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getActiveRange();
  var data = range.getValues();
  var results = [];

  data.forEach(function(row, i) {
    var rowNum = range.getRow() + i;
    var name  = row[0];  // Professor's name
    var email = row[1];  // Professor's email
    var research  = row[2];  // Professor's research focus
    var status = {row: rowNum};

    // 4a. Basic validation
    if (!name || !email) {
      status.error = 'Missing name or email';
      results.push(status);
      return;
    }
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      status.error = 'Invalid email format';
      results.push(status);
      return;
    }

    try {
      // 5. Generate body via OpenAI
      var body = generateEmailBody(name, research);
      // 7. Create draft in Outlook
      createOutlookDraft(email, 'Research Collaboration Inquiry', body);
      status.success = true;

    } catch (e) {
      status.error = e.message || e.toString();
    }
    results.push(status);
  });

  return results;
}
function generateEmailBody(name, firm, note, variant, tone) {
  // Get the professor's last name and add "Dr." prefix
  const lastName = name.split(' ').slice(-1)[0];
  const greeting = `Dr. ${lastName}`;  // "Dr. [Last Name]"

  const research = firm;  // Use the research column for the specific research focus of the professor

  const systemPrompt = `
You are an email generating assistant for a specific purpose '(user enter here)'....
  `.trim();

  const userPrompt = `
Recipient Name: ${greeting}
Research Focus: ${research}

Please craft a personalized email to the professor with the followingc(enter format and description):

1. ...
2. ...
3. ...
4. ...
5. ...
  `.trim();

  const payload = {
    model: 'gpt-4o-mini',
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: userPrompt }
    ],
    temperature: 0.7,
    max_tokens: 350
  };

  const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method      : 'post',
    contentType : 'application/json',
    headers     : { 'Authorization': 'Bearer OPEN AI KEY (<-- ENTER)' },
    payload     : JSON.stringify(payload)
  });

  const text = JSON.parse(response.getContentText())
                .choices[0].message.content
                .trim();

  const paragraphs = text.split(/\n{2,}/).map(p => p.trim());

  // Start with the greeting
  let htmlBody = `<p>Hi ${greeting},</p>`;

  // Add paragraphs and remove extra line breaks
  paragraphs.slice(1).forEach(p => {
    htmlBody += `<p>${p.replace(/\n/g, ' ')}</p>`;
  });

  return htmlBody;
}





/**
 * Creates an Outlook draft via Microsoft Graph
 */
/**
 * Creates an Outlook draft via Microsoft Graph
 */
function createOutlookDraft(toEmail, subject, htmlBody) {
  var service = getOAuthService();
  if (!service.hasAccess()) {
    throw new Error('Authorize this app by visiting:\n' + service.getAuthorizationUrl());
  }
  var token = service.getAccessToken();

  // Build the Message object itself (no "message" wrapper)
  var messagePayload = {
    subject: subject,
    body: {
      contentType: 'HTML',
      content: htmlBody
    },
    toRecipients: [
      { emailAddress: { address: toEmail } }
    ]
  };

  var res = UrlFetchApp.fetch('https://graph.microsoft.com/v1.0/me/messages', {
    method            : 'post',
    contentType       : 'application/json',
    headers           : { Authorization: 'Bearer ' + token },
    payload           : JSON.stringify(messagePayload),
    muteHttpExceptions: true
  });

  var code = res.getResponseCode();
  var text = res.getContentText();
  if (code === 401) {
    service.reset();
    throw new Error('Unauthenticated. Please re-authorize.');
  }
  if (code >= 400) {
    // Log the full response for debugging
    Logger.log('Graph Error: ' + text);
    throw new Error('Graph API error ' + code + ': ' + text);
  }
}


/**
 * Configures OAuth2 for Microsoft Graph
 */
function getOAuthService() {
  return OAuth2.createService('Graph')
    // Standard OAuth2 endpoints for Microsoft
    .setAuthorizationBaseUrl('https://login.microsoftonline.com/common/oauth2/v2.0/authorize')
    .setTokenUrl('https://login.microsoftonline.com/common/oauth2/v2.0/token')
    // Your Azure AD app credentials
    .setClientId('Azure Client ID')
    .setClientSecret('Azure Secret ID')
    // Must match the callback function name in your script
    .setCallbackFunction('authCallback')
    // Persist tokens so users only consent once
    .setPropertyStore(PropertiesService.getUserProperties())
    // Enable caching and locking per best practices
    .setCache(CacheService.getUserCache())
    .setLock(LockService.getUserLock())
    // The scopes you need
    .setScope([
      'https://graph.microsoft.com/Mail.Send',
      'https://graph.microsoft.com/Mail.ReadWrite'
    ].join(' '))
    // Ensure you get a refresh token
    .setParam('access_type', 'offline')
    // Always prompt consent so refresh tokens are returned
    .setParam('prompt', 'consent');
}


/**
 * OAuth2 callback — do not modify the name
 */
function authCallback(request) {
  var service = getOAuthService();
  var authorized = service.handleCallback(request);
  return HtmlService
    .createHtmlOutput( authorized 
      ? 'Success! You may now close this tab.' 
      : 'Denied. You may close this tab.' );
}
