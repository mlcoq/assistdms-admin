// AssistDMS Standalone Web App with Azure AD Authentication
const API_URL = API_CONFIG.baseUrl;

// MSAL Configuration
const msalConfig = {
    auth: {
        clientId: '9635361b-5007-4fa8-8661-c78cb3a1402f',
        authority: 'https://login.microsoftonline.com/e2e6f0bc-a094-4bb3-9250-1975a8102eeb',
        redirectUri: window.location.origin + window.location.pathname
    },
    cache: {
        cacheLocation: 'localStorage',
        storeAuthStateInCookie: false
    }
};

const loginRequest = {
    scopes: [
        'api://9635361b-5007-4fa8-8661-c78cb3a1402f/access_as_user',
        'Mail.Read',
        'Mail.ReadWrite'
    ]
};

const msalInstance = new msal.PublicClientApplication(msalConfig);
let accessToken = null;
let account = null;

// State
let customers = [];
let tickets = [];

// Initialize
document.addEventListener('DOMContentLoaded', async () => {
    await initAuth();
});

// Authentication
async function initAuth() {
    try {
        await msalInstance.initialize();
        console.log('MSAL initialized');

        // Handle redirect response
        const response = await msalInstance.handleRedirectPromise();
        if (response) {
            console.log('Login redirect response:', response);
            account = response.account;
            // Extract access token from redirect response
            if (response.accessToken) {
                accessToken = response.accessToken;
                console.log('Access token obtained from redirect');
            } else {
                console.log('No access token in redirect response, requesting...');
                await getToken();
            }
        } else {
            // Check if user is already logged in
            const accounts = msalInstance.getAllAccounts();
            console.log('Found accounts:', accounts.length);
            if (accounts.length > 0) {
                account = accounts[0];
                console.log('Using existing account:', account.username);
                await getToken();
            } else {
                console.log('No accounts found, showing login');
                showLoginOverlay();
                return;
            }
        }

        // Hide login overlay and show user info
        document.getElementById('login-overlay').style.display = 'none';
        document.getElementById('user-info').textContent = `👤 ${account.name}`;
        console.log('Access token available:', !!accessToken);

        // Load data
        loadCustomers();
        loadTickets();

    } catch (error) {
        console.error('Auth error:', error);
        showLoginOverlay();
    }
}

async function login() {
    try {
        await msalInstance.loginRedirect(loginRequest);
    } catch (error) {
        console.error('Login error:', error);
        document.getElementById('login-error').textContent = 'Login failed: ' + error.message;
        document.getElementById('login-error').style.display = 'block';
    }
}

async function logout() {
    await msalInstance.logoutRedirect({
        account: account
    });
}

async function getToken() {
    console.log('getToken() called, account:', account?.username);

    if (!account) {
        console.error('No account available for token request');
        showLoginOverlay();
        return;
    }

    const tokenRequest = {
        scopes: loginRequest.scopes,
        account: account
    };

    try {
        console.log('Attempting acquireTokenSilent with scopes:', tokenRequest.scopes);
        const response = await msalInstance.acquireTokenSilent(tokenRequest);
        accessToken = response.accessToken;
        console.log('Access token acquired silently, length:', accessToken?.length);
    } catch (error) {
        console.warn('Silent token acquisition failed:', error);

        // Try interactive token acquisition with popup instead of redirect
        try {
            console.log('Attempting acquireTokenPopup...');
            const response = await msalInstance.acquireTokenPopup(tokenRequest);
            accessToken = response.accessToken;
            console.log('Access token acquired via popup, length:', accessToken?.length);
        } catch (popupError) {
            console.error('Popup token acquisition failed:', popupError);
            // If popup also fails, redirect as last resort
            console.log('Falling back to redirect...');
            await msalInstance.acquireTokenRedirect(tokenRequest);
        }
    }
}

// Get token specifically for Microsoft Graph API
async function getGraphToken() {
    console.log('getGraphToken() called');

    if (!account) {
        console.error('No account available for Graph token request');
        return null;
    }

    const graphTokenRequest = {
        scopes: ['Mail.Read', 'Mail.ReadWrite'],
        account: account
    };

    try {
        console.log('Attempting acquireTokenSilent for Graph API...');
        const response = await msalInstance.acquireTokenSilent(graphTokenRequest);
        console.log('Graph token acquired silently, length:', response.accessToken?.length);
        return response.accessToken;
    } catch (error) {
        console.warn('Silent Graph token acquisition failed:', error);

        // Try interactive token acquisition with popup
        try {
            console.log('Attempting acquireTokenPopup for Graph API...');
            const response = await msalInstance.acquireTokenPopup(graphTokenRequest);
            console.log('Graph token acquired via popup, length:', response.accessToken?.length);
            return response.accessToken;
        } catch (popupError) {
            console.error('Popup Graph token acquisition failed:', popupError);
            throw popupError;
        }
    }
}

function showLoginOverlay() {
    document.getElementById('login-overlay').style.display = 'flex';
}

// Authenticated fetch wrapper
async function authFetch(url, options = {}) {
    if (!accessToken) {
        console.log('No access token, requesting...');
        await getToken();
    }

    if (!accessToken) {
        console.error('Failed to obtain access token for API call');
        throw new Error('Not authenticated - no access token available');
    }

    const headers = {
        ...options.headers,
        'Authorization': `Bearer ${accessToken}`
    };

    console.log('Making authenticated request to:', url);
    const response = await fetch(url, { ...options, headers });

    if (response.status === 401) {
        console.warn('401 Unauthorized, token may be expired, requesting new token...');
        // Token expired, get new one
        await getToken();

        if (!accessToken) {
            console.error('Failed to refresh access token');
            throw new Error('Authentication failed - unable to refresh token');
        }

        headers['Authorization'] = `Bearer ${accessToken}`;
        console.log('Retrying request with new token...');
        return fetch(url, { ...options, headers });
    }

    return response;
}

// Tab switching
function showTab(tabName, event) {
    document.querySelectorAll('.tab').forEach(tab => tab.classList.remove('active'));
    document.querySelectorAll('.content').forEach(content => content.classList.remove('active'));

    if (event && event.target) {
        event.target.classList.add('active');
    }
    document.getElementById(tabName).classList.add('active');

    // Clear status messages
    document.querySelectorAll('.status').forEach(s => s.style.display = 'none');

    // Load emails when switching to emails tab
    if (tabName === 'emails') {
        loadEmails();
    }
}

// Load customers
async function loadCustomers() {
    try {
        const response = await authFetch(`${API_URL}/customers`);
        customers = await response.json();
        
        renderCustomers();
        populateCustomerSelect();
    } catch (error) {
        showStatus('customers', `Fout bij laden klanten: ${error.message}`, 'error');
    }
}

// Load tickets
async function loadTickets() {
    try {
        const response = await authFetch(`${API_URL}/tickets`);
        tickets = await response.json();
        
        renderTickets();
        populateTicketSelect();
    } catch (error) {
        document.getElementById('tickets-list').innerHTML = 
            `<div class="empty-state">❌ Fout bij laden: ${error.message}</div>`;
    }
}

// Render customers
function renderCustomers() {
    const list = document.getElementById('customers-list');

    if (customers.length === 0) {
        list.innerHTML = '<div class="empty-state">📭 Nog geen klanten</div>';
        return;
    }

    list.innerHTML = customers.map(customer => `
        <div class="card">
            <h3>${customer.name}</h3>
            ${customer.defaultArchiveMailbox ? `<p>📂 Archief: ${customer.defaultArchiveMailbox}</p>` : ''}
        </div>
    `).join('');
}

// Render tickets
function renderTickets() {
    const list = document.getElementById('tickets-list');
    
    if (tickets.length === 0) {
        list.innerHTML = '<div class="empty-state">📭 Nog geen tickets</div>';
        return;
    }
    
    list.innerHTML = tickets.map(ticket => {
        const customer = customers.find(c => c.id === ticket.customerId);
        const statusClass = ticket.status === 'Open' ? 'open' : 
                          ticket.status === 'Closed' ? 'closed' : 'archived';
        
        return `
            <div class="card">
                <h3>${ticket.subject}</h3>
                <p><strong>Klant:</strong> ${customer ? customer.name : 'Onbekend'}</p>
                <p><strong>Aangemaakt:</strong> ${new Date(ticket.createdAt).toLocaleDateString('nl-NL')}</p>
                <p>
                    <span class="badge ${statusClass}">${ticket.status}</span>
                    ${ticket.mailsCount > 0 ? `<span class="badge" style="background:#007bff;color:white">📧 ${ticket.mailsCount} emails</span>` : ''}
                    ${ticket.totalHours > 0 ? `<span class="badge" style="background:#ffc107;color:#000">⏱️ ${ticket.totalHours}u</span>` : ''}
                </p>
                ${ticket.description ? `<p style="color:#666;font-size:14px">${ticket.description}</p>` : ''}
            </div>
        `;
    }).join('');
}

// Populate selects
function populateCustomerSelect() {
    const select = document.getElementById('customer-select');
    select.innerHTML = '<option value="">Selecteer klant...</option>' +
        customers.map(c => `<option value="${c.id}">${c.name}</option>`).join('');
}

function populateTicketSelect() {
    const select = document.getElementById('ticket-select');
    const openTickets = tickets.filter(t => t.status === 'Open');
    
    select.innerHTML = '<option value="">Selecteer ticket...</option>' +
        openTickets.map(t => `<option value="${t.id}">${t.subject}</option>`).join('');
}

// Create customer
async function createCustomer() {
    const name = document.getElementById('customer-name').value.trim();
    const archive = document.getElementById('customer-archive').value.trim();

    if (!name) {
        showStatus('customers', 'Naam is verplicht', 'error');
        return;
    }
    
    try {
        const response = await authFetch(`${API_URL}/customers`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                Name: name,
                DefaultArchiveMailbox: archive || null
            })
        });

        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`API error: ${errorText}`);
        }
        
        showStatus('customers', `✅ Klant "${name}" aangemaakt!`, 'success');

        // Clear form
        document.getElementById('customer-name').value = '';
        document.getElementById('customer-archive').value = '';
        
        // Reload
        await loadCustomers();
        
    } catch (error) {
        showStatus('customers', `❌ Fout: ${error.message}`, 'error');
    }
}

// Create ticket
async function createTicket() {
    const customerId = document.getElementById('customer-select').value;
    const subject = document.getElementById('ticket-subject').value.trim();
    const description = document.getElementById('ticket-description').value.trim();

    const emailSubject = document.getElementById('email-subject').value.trim();
    const emailId = document.getElementById('email-id').value.trim();
    
    if (!customerId || !subject) {
        showStatus('new-ticket', 'Klant en onderwerp zijn verplicht', 'error');
        return;
    }
    
    try {
        // Create ticket
        const ticketResponse = await authFetch(`${API_URL}/tickets`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                CustomerId: customerId,
                Title: subject,
                ArchiveMailboxOverride: null
            })
        });

        if (!ticketResponse.ok) {
            const errorText = await ticketResponse.text();
            throw new Error(`Fout bij aanmaken ticket: ${errorText}`);
        }
        
        const ticket = await ticketResponse.json();

        // Link email if provided
        if (emailSubject) {
            const mailResponse = await authFetch(`${API_URL}/tickets/${ticket.id}/mails`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    Subject: emailSubject,
                    OutlookMessageId: emailId || null
                })
            });

            if (!mailResponse.ok) {
                console.warn('Email koppelen mislukt, maar ticket is aangemaakt');
            }
        }
        
        showStatus('new-ticket', `✅ Ticket "${subject}" aangemaakt!`, 'success');

        // Clear form
        document.getElementById('customer-select').value = '';
        document.getElementById('ticket-subject').value = '';
        document.getElementById('ticket-description').value = '';
        document.getElementById('email-subject').value = '';
        document.getElementById('email-id').value = '';
        
        // Reload tickets
        await loadTickets();
        
        // Switch to tickets tab
        setTimeout(() => {
            document.querySelector('.tab').click();
        }, 2000);
        
    } catch (error) {
        showStatus('new-ticket', `❌ Fout: ${error.message}`, 'error');
    }
}

// Link email
async function linkEmail() {
    const ticketId = document.getElementById('ticket-select').value;
    const subject = document.getElementById('link-email-subject').value.trim();
    const messageId = document.getElementById('link-email-id').value.trim();

    if (!ticketId || !subject) {
        showStatus('email', 'Ticket en onderwerp zijn verplicht', 'error');
        return;
    }
    
    try {
        const response = await authFetch(`${API_URL}/tickets/${ticketId}/mails`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                Subject: subject,
                OutlookMessageId: messageId || null
            })
        });

        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`API error: ${errorText}`);
        }
        
        showStatus('email', '✅ Email gekoppeld aan ticket!', 'success');

        // Clear form
        document.getElementById('ticket-select').value = '';
        document.getElementById('link-email-subject').value = '';
        document.getElementById('link-email-id').value = '';
        
        await loadTickets();
        
    } catch (error) {
        showStatus('email', `❌ Fout: ${error.message}`, 'error');
    }
}

// Archive email (link + archive)
async function archiveEmail() {
    const ticketId = document.getElementById('ticket-select').value;
    const subject = document.getElementById('link-email-subject').value.trim();
    const from = document.getElementById('link-email-from').value.trim();
    
    if (!ticketId || !subject || !from) {
        showStatus('email', 'Ticket, onderwerp en afzender zijn verplicht', 'error');
        return;
    }
    
    try {
        // Link email first
        await linkEmail();
        
        // Then mark as archived (simplified - in reality you'd get the mail ID back)
        showStatus('email', '✅ Email gekoppeld en gearchiveerd!', 'success');
        
    } catch (error) {
        showStatus('email', `❌ Fout: ${error.message}`, 'error');
    }
}

// Show status message
function showStatus(section, message, type) {
    const statusEl = document.getElementById(`status-${section}`);
    statusEl.className = `status ${type}`;
    statusEl.textContent = message;
    statusEl.style.display = 'block';

    // Auto-hide success messages
    if (type === 'success') {
        setTimeout(() => {
            statusEl.style.display = 'none';
        }, 5000);
    }
}

// ===== EMAIL BROWSER FUNCTIONALITY =====

const ASSISTDMS_CATEGORY = 'AssistDMS Gekoppeld';
let emails = [];

// Load emails from Outlook via Microsoft Graph
async function loadEmails() {
    const showLinked = document.getElementById('show-linked-emails').checked;

    try {
        document.getElementById('emails-list').innerHTML = '<div class="loading">Emails laden...</div>';

        // Get token specifically for Graph API
        const graphToken = await getGraphToken();

        if (!graphToken) {
            throw new Error('Kon geen Graph API token verkrijgen');
        }

        console.log('Calling Graph API with token...');

        // Call Microsoft Graph API to get emails
        const graphResponse = await fetch('https://graph.microsoft.com/v1.0/me/messages?$top=50&$orderby=receivedDateTime desc&$select=id,subject,from,receivedDateTime,bodyPreview,categories,internetMessageId', {
            headers: {
                'Authorization': `Bearer ${graphToken}`
            }
        });

        if (!graphResponse.ok) {
            const errorText = await graphResponse.text();
            console.error('Graph API error response:', errorText);
            throw new Error(`Graph API error: ${graphResponse.status} - ${errorText}`);
        }

        const data = await graphResponse.json();
        emails = data.value;

        console.log(`Loaded ${emails.length} emails from Graph API`);

        // Filter based on checkbox
        let displayEmails = emails;
        if (!showLinked) {
            displayEmails = emails.filter(email => !email.categories || !email.categories.includes(ASSISTDMS_CATEGORY));
        }

        renderEmails(displayEmails);

    } catch (error) {
        console.error('Load emails error:', error);
        document.getElementById('emails-list').innerHTML = 
            `<div class="empty-state">❌ Fout bij laden emails: ${error.message}</div>`;
    }
}

// Render emails list
function renderEmails(emailsList) {
    const list = document.getElementById('emails-list');

    if (emailsList.length === 0) {
        list.innerHTML = '<div class="empty-state">📭 Geen emails gevonden</div>';
        return;
    }

    list.innerHTML = emailsList.map(email => {
        const isLinked = email.categories && email.categories.includes(ASSISTDMS_CATEGORY);
        const fromEmail = email.from?.emailAddress?.address || 'Onbekend';
        const fromName = email.from?.emailAddress?.name || fromEmail;
        const date = new Date(email.receivedDateTime).toLocaleString('nl-NL');

        return `
            <div class="card" style="border-left-color: ${isLinked ? '#28a745' : '#667eea'}">
                <div style="display:flex; justify-content:space-between; align-items:start;">
                    <div style="flex:1;">
                        <h3>${email.subject || '(Geen onderwerp)'}</h3>
                        <p><strong>Van:</strong> ${fromName} &lt;${fromEmail}&gt;</p>
                        <p><strong>Datum:</strong> ${date}</p>
                        ${email.bodyPreview ? `<p style="color:#999; font-size:13px; margin-top:10px;">${email.bodyPreview.substring(0, 150)}...</p>` : ''}
                        ${isLinked ? '<p style="color:#28a745; margin-top:10px;">✅ Gekoppeld aan ticket</p>' : ''}
                    </div>
                    ${!isLinked ? `
                        <button class="btn" onclick="showLinkEmailModal('${email.id}')" style="margin-left:15px;">
                            Koppel aan Ticket
                        </button>
                    ` : ''}
                </div>
            </div>
        `;
    }).join('');
}

// Show modal to link email to ticket
async function showLinkEmailModal(emailId) {
    const email = emails.find(e => e.id === emailId);
    if (!email) return;

    // Pre-fill the "Email Koppelen" tab
    document.getElementById('link-email-subject').value = email.subject || '';
    document.getElementById('link-email-id').value = email.internetMessageId || '';

    // Store email ID for later use
    window.currentEmailId = emailId;

    // Switch to Email Koppelen tab
    showTab('email');
    document.querySelector('.tab:nth-child(5)').click();

    showStatus('email', '💡 Email info ingevuld - selecteer een ticket en klik "Email Koppelen"', 'success');
}

// Modified linkEmail to add category
const originalLinkEmail = linkEmail;
async function linkEmail() {
    try {
        // Call original linkEmail function
        await originalLinkEmail();

        // If we have a currentEmailId, add category
        if (window.currentEmailId) {
            await addCategoryToEmail(window.currentEmailId);
            window.currentEmailId = null;
        }

    } catch (error) {
        throw error;
    }
}

// Add AssistDMS category to email
async function addCategoryToEmail(emailId) {
    try {
        const graphToken = await getGraphToken();

        if (!graphToken) {
            console.error('No Graph token available');
            return;
        }

        // Get current categories
        const email = emails.find(e => e.id === emailId);
        const currentCategories = email?.categories || [];

        // Add our category if not already present
        if (!currentCategories.includes(ASSISTDMS_CATEGORY)) {
            const newCategories = [...currentCategories, ASSISTDMS_CATEGORY];

            console.log('Adding category to email:', emailId);

            // Update via Graph API
            const response = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${emailId}`, {
                method: 'PATCH',
                headers: {
                    'Authorization': `Bearer ${graphToken}`,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    categories: newCategories
                })
            });

            if (!response.ok) {
                const errorText = await response.text();
                console.error('Failed to add category:', errorText);
            } else {
                console.log('✅ Category "' + ASSISTDMS_CATEGORY + '" added to email in Outlook');
            }
        }

    } catch (error) {
        console.error('Add category error:', error);
    }
}
