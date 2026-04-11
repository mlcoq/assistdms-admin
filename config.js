// AssistDMS API Configuration
// Wijzig deze URL naar waar je API draait

const API_CONFIG = {
    // Production API (Railway hosting):
    baseUrl: 'https://assistdms-production.up.railway.app'

    // Voor lokaal testen (alleen op je eigen machine):
    // baseUrl: 'http://localhost:5001'

    // Of gebruik Ngrok voor tijdelijk testen:
    // baseUrl: 'https://abc123.ngrok.io'
};

// Gebruik deze in je code als: API_CONFIG.baseUrl
