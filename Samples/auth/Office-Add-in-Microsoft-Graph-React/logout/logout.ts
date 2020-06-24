/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import * as msal from 'msal';

(() => {
  // The initialize function must be run each time a new page is loaded
  Office.initialize = () => {

    const config: msal.Configuration = {
      auth: {
          clientId: 'd2dbd3e1-9f0f-44ef-a975-a3f4cab6e4ff',
          redirectUri: 'https://emission-factors.azurewebsites.net/logoutcomplete/logoutcomplete.html', 
          postLogoutRedirectUri: 'https://emission-factors.azurewebsites.net/logoutcomplete/logoutcomplete.html'
      }
    };

    const userAgentApplication = new msal.UserAgentApplication(config);
    userAgentApplication.logout();
  };
})();
