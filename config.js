// AssistDMS API Configuration
// Wijzig deze URL naar waar je API draait

const API_CONFIG = {
    // Voor lokaal testen (alleen op je eigen machine):
    // baseUrl: 'https://localhost:5001'
    
    // Voor production (vervang met je echte API URL):
    baseUrl: 'https://jouw-api-url.azurewebsites.net'
    
    // Of gebruik Ngrok voor tijdelijk testen:
    // baseUrl: 'https://abc123.ngrok.io'
};

// Gebruik deze in je code als: API_CONFIG.baseUrl
