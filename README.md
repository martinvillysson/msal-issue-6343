# Msal-Issue-6343
https://github.com/AzureAD/microsoft-authentication-library-for-js/issues/6343

This project was generated with [Angular CLI](https://github.com/angular/angular-cli) version 16.2.0.

## Development server

Run `ng serve` for a dev server. Navigate to `http://localhost:4200/`. The application will automatically reload if you change any of the source files.

## Instructions
### Ensure you have a Azure AD app registration
1. Ensure the app registration has correct redirect uri's and permissions for Microsoft Graph, Yammer and SharePoint.
1. Replace URLs for services and or scopes in `environment.ts` and `app.module.ts` (check the REPLACE_ME entries)
1. Fill in the client id and tenant id in `environment.ts`.
1. Run `npm start`

### Open devtools F12 and Check the console!

1. Make sure the page is loaded and you got entries in the console
1. Delete access token from localStorage
1. Modify the refresh token entry in localStorage by removing a letter from the "secret" entry
1. Reload page and it should start looping
