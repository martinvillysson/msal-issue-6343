import { APP_INITIALIZER, NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';
import { AppRoutingModule } from './app-routing.module';
import { AppComponent } from './app.component';
import { firstValueFrom } from 'rxjs';
import { MsalBroadcastService, MsalInterceptor, MsalModule, MsalService } from '@azure/msal-angular';
import { InteractionType, PublicClientApplication } from '@azure/msal-browser';
import { HTTP_INTERCEPTORS, HttpClientModule } from '@angular/common/http';
import { environment } from './environment';

export function initAndAuthenticate(msal: MsalService): { (): Promise<any> } {
  // abort if not in a browser or if in iframe
  if (window.self !== window.top) {
    return () => Promise.resolve(null);
  }

  return () => firstValueFrom(msal.handleRedirectObservable());
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
