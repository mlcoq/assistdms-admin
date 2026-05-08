// AssistDMS API Configuration
// Wijzig deze URL naar waar je API draait

const API_CONFIG = {
    // Production API (Railway hosting):
    baseUrl: 'https://assistdms-production.up.railway.app',

    // Externe tijdschrijfapp (optioneel, opent in nieuwe tab)
    // Voorbeeld: 'https://tijdschrijven.jouwdomein.nl'
    timeWriterUrl: '',

    // TimeWriter aspectType used to import customer list
    // Veel installaties gebruiken IT_AT1 als klantdimension
    timeWriterCustomerAspectType: 'IT_AT1'

    // Voor lokaal testen (alleen op je eigen machine):
    // baseUrl: 'http://localhost:5001'

    // Of gebruik Ngrok voor tijdelijk testen:
    // baseUrl: 'https://abc123.ngrok.io'
};

// Gebruik deze in je code als: API_CONFIG.baseUrl
