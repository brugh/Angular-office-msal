import { enableProdMode } from '@angular/core';
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';

import { AppModule } from './app/app.module';
import { environment } from './environments/environment';

declare const Office: any;

if (environment.production) {
  enableProdMode();
}
(async () => {
  await Office.onReady();
  // if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
  //   console.log("Sorry, this add-in only works with newer versions of Excel.");
  // }
  platformBrowserDynamic().bootstrapModule(AppModule)
    .catch(err => console.error(err));
})();
