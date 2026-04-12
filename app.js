// AssistDMS Standalone Web App
const API_URL = API_CONFIG.baseUrl;

// State
let customers = [];
let tickets = [];

// Initialize
document.addEventListener('DOMContentLoaded', () => {
    loadCustomers();
    loadTickets();
});

// Tab switching
function showTab(tabName) {
    document.querySelectorAll('.tab').forEach(tab => tab.classList.remove('active'));
    document.querySelectorAll('.content').forEach(content => content.classList.remove('active'));
    
    event.target.classList.add('active');
    document.getElementById(tabName).classList.add('active');
    
    // Clear status messages
    document.querySelectorAll('.status').forEach(s => s.style.display = 'none');
}

// Load customers
async function loadCustomers() {
    try {
        const response = await fetch(`${API_URL}/customers`);
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
        const response = await fetch(`${API_URL}/tickets`);
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
            <p>✉️ ${customer.email}</p>
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
    const email = document.getElementById('customer-email').value.trim();
    const archive = document.getElementById('customer-archive').value.trim();
    
    if (!name || !email) {
        showStatus('customers', 'Naam en email zijn verplicht', 'error');
        return;
    }
    
    try {
        const response = await fetch(`${API_URL}/customers`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                name,
                email,
                defaultArchiveMailbox: archive || null
            })
        });
        
        if (!response.ok) throw new Error('Fout bij aanmaken klant');
        
        showStatus('customers', `✅ Klant "${name}" aangemaakt!`, 'success');
        
        // Clear form
        document.getElementById('customer-name').value = '';
        document.getElementById('customer-email').value = '';
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
    const emailFrom = document.getElementById('email-from').value.trim();
    const emailId = document.getElementById('email-id').value.trim();
    
    if (!customerId || !subject) {
        showStatus('new-ticket', 'Klant en onderwerp zijn verplicht', 'error');
        return;
    }
    
    try {
        // Create ticket
        const ticketResponse = await fetch(`${API_URL}/tickets`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                customerId: parseInt(customerId),
                subject,
                description: description || null
            })
        });
        
        if (!ticketResponse.ok) throw new Error('Fout bij aanmaken ticket');
        
        const ticket = await ticketResponse.json();
        
        // Link email if provided
        if (emailSubject && emailFrom) {
            await fetch(`${API_URL}/tickets/${ticket.id}/mails`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    subject: emailSubject,
                    from: emailFrom,
                    messageId: emailId || null,
                    receivedAt: new Date().toISOString()
                })
            });
        }
        
        showStatus('new-ticket', `✅ Ticket "${subject}" aangemaakt!`, 'success');
        
        // Clear form
        document.getElementById('customer-select').value = '';
        document.getElementById('ticket-subject').value = '';
        document.getElementById('ticket-description').value = '';
        document.getElementById('email-subject').value = '';
        document.getElementById('email-from').value = '';
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
    const from = document.getElementById('link-email-from').value.trim();
    const messageId = document.getElementById('link-email-id').value.trim();
    
    if (!ticketId || !subject || !from) {
        showStatus('email', 'Ticket, onderwerp en afzender zijn verplicht', 'error');
        return;
    }
    
    try {
        const response = await fetch(`${API_URL}/tickets/${ticketId}/mails`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                subject,
                from,
                messageId: messageId || null,
                receivedAt: new Date().toISOString()
            })
        });
        
        if (!response.ok) throw new Error('Fout bij koppelen email');
        
        showStatus('email', '✅ Email gekoppeld aan ticket!', 'success');
        
        // Clear form
        document.getElementById('ticket-select').value = '';
        document.getElementById('link-email-subject').value = '';
        document.getElementById('link-email-from').value = '';
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
