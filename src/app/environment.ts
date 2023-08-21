import { BrowserCacheLocation, Configuration, LogLevel } from "@azure/msal-browser";

export interface MsalConfiguration extends Configuration {
  protectedResourceScopeMap?: [string, string[]][] | Map<string, string[]>;
}

export class Environment {
  /** MSAL login information */

  public msalConfig:Configuration = {
    auth: {
      clientId: 'CLIENT_ID_REPLACE_ME',
      redirectUri: 'http://localhost:4200',
      authority: 'https://login.microsoftonline.com/TENANT_ID_REPLACE_ME',
      navigateToLoginRequestUrl: true,
    },
    cache: {
      cacheLocation: BrowserCacheLocation.LocalStorage,
    },
    system: {
      loggerOptions: {
        logLevel: LogLevel.Verbose,
        loggerCallback: (level, message) => { console.log(message) }
      },
    },
  }

  public endpoints= {
    sharepoint: 'https://REPLACE_ME.sharepoint.com'
  };
}

export const environment = new Environment();
