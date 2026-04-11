// API Base URL - gebruikt config.js
const API_BASE = API_CONFIG.baseUrl;

let currentEmailItem = null;
let currentEmailSubject = '';
let currentEmailFrom = '';
let currentMessageId = '';

// Initialize Office
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        currentEmailItem = Office.context.mailbox.item;
        loadEmailInfo();
        loadCustomers();
    }
});

// Load current email information
function loadEmailInfo() {
    if (!currentEmailItem) return;
    
    currentEmailItem.subject.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            currentEmailSubject = result.value;
            document.getElementById('emailSubject').textContent = currentEmailSubject;
        }
    });
    
    currentEmailItem.from.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            currentEmailFrom = result.value.displayName + ' <' + result.value.emailAddress + '>';
            document.getElementById('emailFrom').textContent = currentEmailFrom;
        }
    });
    
    currentEmailItem.internetMessageId.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            currentMessageId = result.value;
        }
    });
}

// API Helper Functions
async function apiCall(endpoint, method = 'GET', body = null) {
    const options = {
        method: method,
        headers: {
            'Content-Type': 'application/json',
        }
    };
    
    if (body) {
        options.body = JSON.stringify(body);
    }
    
    try {
        const response = await fetch(`${API_BASE}${endpoint}`, options);
        
        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(errorText || `HTTP ${response.status}`);
        }
        
        return await response.json();
    } catch (error) {
        console.error('API Error:', error);
        throw error;
    }
}

// Load customers
async function loadCustomers() {
    try {
        const customers = await apiCall('/customers');
        
        const customerSelect = document.getElementById('customerSelect');
        const archiveTicketSelect = document.getElementById('archiveTicketSelect');
        const timeTicketSelect = document.getElementById('timeTicketSelect');
        
        customerSelect.innerHTML = '<option value="">Selecteer klant...</option>';
        
        customers.forEach(customer => {
            const option = document.createElement('option');
            option.value = customer.id;
            option.textContent = customer.name;
            customerSelect.appendChild(option);
        });
    } catch (error) {
        showStatus('Fout bij laden van klanten: ' + error.message, 'error');
    }
}

// Load tickets
async function loadTickets() {
    try {
        const tickets = await apiCall('/tickets');
        
        const ticketList = document.getElementById('ticketList');
        ticketList.innerHTML = '';
        
        if (tickets.length === 0) {
            ticketList.innerHTML = '<div class="loading">Geen tickets gevonden</div>';
            return;
        }
        
        tickets.forEach(ticket => {
            const item = document.createElement('div');
            item.className = 'ticket-item';
            item.innerHTML = `
                <div class="ticket-number">${ticket.number}</div>
                <div>${ticket.title}</div>
                <small>${ticket.status}</small>
            `;
            item.onclick = () => linkEmailToTicket(ticket.id);
            ticketList.appendChild(item);
        });
        
        // Also populate select dropdowns
        const archiveSelect = document.getElementById('archiveTicketSelect');
        const timeSelect = document.getElementById('timeTicketSelect');
        
        archiveSelect.innerHTML = '<option value="">Selecteer ticket...</option>';
        timeSelect.innerHTML = '<option value="">Selecteer ticket...</option>';
        
        tickets.forEach(ticket => {
            const option1 = document.createElement('option');
            option1.value = ticket.id;
            option1.textContent = `${ticket.number} - ${ticket.title}`;
            archiveSelect.appendChild(option1);
            
            const option2 = document.createElement('option');
            option2.value = ticket.id;
            option2.textContent = `${ticket.number} - ${ticket.title}`;
            timeSelect.appendChild(option2);
        });
        
    } catch (error) {
        showStatus('Fout bij laden van tickets: ' + error.message, 'error');
    }
}

// Show forms
function showNewTicketForm() {
    hideAllForms();
    document.getElementById('newTicketForm').classList.remove('hidden');
    document.getElementById('ticketTitle').value = currentEmailSubject;
}

function showLinkToTicketForm() {
    hideAllForms();
    document.getElementById('linkTicketForm').classList.remove('hidden');
    loadTickets();
}

function showArchiveForm() {
    hideAllForms();
    document.getElementById('archiveForm').classList.remove('hidden');
    loadTickets();
}

function showTimeEntryForm() {
    hideAllForms();
    document.getElementById('timeEntryForm').classList.remove('hidden');
    loadTickets();
}

function hideAllForms() {
    document.querySelectorAll('.section:not(:nth-child(1)):not(:nth-child(2))').forEach(section => {
        if (!section.querySelector('h3').textContent.includes('Acties')) {
            section.classList.add('hidden');
        }
    });
}

// Create new ticket
async function createNewTicket() {
    const customerId = document.getElementById('customerSelect').value;
    const title = document.getElementById('ticketTitle').value;
    const archiveMailbox = document.getElementById('archiveMailbox').value;
    
    if (!customerId || !title) {
        showStatus('Selecteer een klant en voer een titel in', 'error');
        return;
    }
    
    try {
        const ticket = await apiCall('/tickets', 'POST', {
            customerId: customerId,
            title: title,
            archiveMailboxOverride: archiveMailbox || null
        });
        
        showStatus(`✅ Ticket ${ticket.number} aangemaakt!`, 'success');
        
        // Link email to the new ticket
        await linkEmailToTicket(ticket.id);
        
        hideAllForms();
    } catch (error) {
        showStatus('Fout bij aanmaken ticket: ' + error.message, 'error');
    }
}

// Link email to ticket
async function linkEmailToTicket(ticketId) {
    try {
        const mail = await apiCall(`/tickets/${ticketId}/mails`, 'POST', {
            subject: currentEmailSubject,
            outlookMessageId: currentMessageId
        });
        
        showStatus(`✅ Email gekoppeld aan ticket!`, 'success');
        hideAllForms();
    } catch (error) {
        showStatus('Fout bij koppelen email: ' + error.message, 'error');
    }
}

// Archive email
async function archiveEmail() {
    const ticketId = document.getElementById('archiveTicketSelect').value;
    const archiveMailbox = document.getElementById('archiveMailboxOverride').value;
    
    if (!ticketId) {
        showStatus('Selecteer een ticket', 'error');
        return;
    }
    
    try {
        // First, link the email if not already linked
        const ticket = await apiCall(`/tickets/${ticketId}`);
        
        // Find if email is already linked
        let mailId = null;
        for (const mail of ticket.mails) {
            if (mail.outlookMessageId === currentMessageId || mail.subject === currentEmailSubject) {
                mailId = mail.id;
                break;
            }
        }
        
        // If not linked, link it first
        if (!mailId) {
            const mail = await apiCall(`/tickets/${ticketId}/mails`, 'POST', {
                subject: currentEmailSubject,
                outlookMessageId: currentMessageId
            });
            mailId = mail.id;
        }
        
        // Now archive it
        await apiCall(`/tickets/${ticketId}/mails/${mailId}/archive`, 'POST', {
            archiveMailbox: archiveMailbox || null
        });
        
        showStatus(`✅ Email gearchiveerd!`, 'success');
        hideAllForms();
    } catch (error) {
        showStatus('Fout bij archiveren: ' + error.message, 'error');
    }
}

// Add time entry
async function addTimeEntry() {
    const ticketId = document.getElementById('timeTicketSelect').value;
    const description = document.getElementById('timeDescription').value;
    const hours = parseFloat(document.getElementById('timeHours').value);
    
    if (!ticketId || !description || !hours) {
        showStatus('Vul alle velden in', 'error');
        return;
    }
    
    try {
        const now = new Date();
        const startTime = new Date(now.getTime() - (hours * 60 * 60 * 1000));
        
        const entry = await apiCall(`/tickets/${ticketId}/time-entries`, 'POST', {
            description: description,
            startAt: startTime.toISOString(),
            endAt: now.toISOString()
        });
        
        showStatus(`✅ ${entry.billableHoursRounded} uur geschreven (afgerond van ${entry.originalHours.toFixed(2)})`, 'success');
        hideAllForms();
        
        // Clear form
        document.getElementById('timeDescription').value = '';
        document.getElementById('timeHours').value = '';
    } catch (error) {
        showStatus('Fout bij schrijven uren: ' + error.message, 'error');
    }
}

// Show status message
function showStatus(message, type = 'info') {
    const statusDiv = document.getElementById('status');
    statusDiv.textContent = message;
    statusDiv.className = `status ${type}`;
    statusDiv.style.display = 'block';
    
    // Auto-hide after 5 seconds
    setTimeout(() => {
        statusDiv.style.display = 'none';
    }, 5000);
}
