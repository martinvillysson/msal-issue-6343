import { APP_INITIALIZER, NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';
import { AppRoutingModule } from './app-routing.module';
import { AppComponent } from './app.component';
import { catchError, firstValueFrom, map, mergeMap, of } from 'rxjs';
import { MsalBroadcastService, MsalInterceptor, MsalModule, MsalService } from '@azure/msal-angular';
import { InteractionType, PublicClientApplication } from '@azure/msal-browser';
import { HTTP_INTERCEPTORS, HttpClientModule } from '@angular/common/http';
import { environment } from './environment';

export function initAndAuthenticate(msal: MsalService): { (): Promise<any> } {
  // abort if not in a browser or if in iframe
  if (window.self !== window.top) {
    return () => Promise.resolve(null);
  }

  return () => firstValueFrom(msal.handleRedirectObservable().pipe(
    mergeMap((result) => {
      if (result) {
        // If we have a result it means we're coming back from a redirect and can continue
        msal.instance.setActiveAccount(result.account);
        return of(true);
      }

      const account = msal.instance.getActiveAccount() || msal.instance.getAllAccounts()?.[0];
      if (account) {
        /** If we have a stored account in the cache we can try to trigger a silent flow.
         * This ensures the following:
         * 1. A valid access token in cache
         *   a.) If we do NOT have a valid access token it will use the refresh token to renew the access token
         * 2. If the refresh token is invalid or expired the silent flow will fail with code: "interaction_needed" in which case we trigger the loginRedirect which will re-initiate the tokens
        */
        return msal.acquireTokenSilent(
          {
            account,
            // Default scopes used for login (https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/resources-and-scopes.md?plain=1#L16C22-L16C22)
            scopes: ['openid', 'profile', 'offline_access'],
          },
        ).pipe(
          // If we fail we need to trigger interaction and a full redirect
          catchError(() => msal.loginRedirect().pipe(map(() => true))),
          // If we succeed we can continue
          map(() => true),
        );
      }

      // If we have nothing of the above we trigger the normal login flow
      return msal.loginRedirect().pipe(map(() => true));
    }),
    catchError(ex => {

      const { log } = console;
      log('MSAL ERROR', ex);

      return of(false);
    })));
}

@NgModule({
  declarations: [
    AppComponent
  ],
  imports: [
    BrowserModule,
    AppRoutingModule,
    HttpClientModule,
    // authentication
    MsalModule.forRoot(new PublicClientApplication({
      ...environment.msalConfig,
      auth: { ...environment.msalConfig.auth },
      cache: { ...environment.msalConfig.cache },
    }),

      // guard config
      { interactionType: InteractionType.Redirect },

      // interceptor config
      {
        interactionType: InteractionType.Redirect,
        protectedResourceMap: new Map([
          ['https://graph.microsoft.com', ['https://graph.microsoft.com/.default']],
          ['https://api.yammer.com/api/v1', ['https://www.yammer.com/.default']],
          [`${environment.endpoints.sharepoint}`, ['https://REPLACE_ME.sharepoint.com/.default']],
        ]),
      }),

  ],
  providers: [
    { provide: APP_INITIALIZER, useFactory: initAndAuthenticate, multi: true, deps: [MsalService, MsalBroadcastService] },
    { provide: HTTP_INTERCEPTORS, useClass: MsalInterceptor, multi: true },
  ],
  bootstrap: [AppComponent]
})
export class AppModule { }
