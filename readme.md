# SharePoint Contact Sync

Automatische Synchronisation von SharePoint-Listen zu Microsoft Teams Kontakten.

## 🚀 Ein-Klick Deployment

[![Deploy to Azure](https://aka.ms/deploytoazurebutton)](https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fraw.githubusercontent.com%2FIhrUsername%2Fsharepoint-contact-sync%2Fmain%2Fazuredeploy.json)

## Was wird deployed?

- ✅ Azure Function App (PowerShell 7.4)
- ✅ Key Vault für sichere Konfiguration  
- ✅ Application Insights für Monitoring
- ✅ Alle notwendigen Berechtigungen
- ✅ Sofort einsatzbereit

## Nach dem Deployment

1. **Function testen**: `https://ihre-function-app.azurewebsites.net/api/SharePointSync`
2. **Azure AD App Registration** für SharePoint-Zugriff erstellen
3. **SharePoint Berechtigungen** konfigurieren
4. **Teams Kontakte** synchronisieren

## Support

Bei Fragen: [Issues erstellen](https://github.com/IhrUsername/sharepoint-contact-sync/issues)
