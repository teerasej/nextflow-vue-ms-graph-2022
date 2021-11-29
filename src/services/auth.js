// ----------------------------------------------------------------------------
// Copyright (c) Ben Coleman, 2021
// Licensed under the MIT License.
//
// Drop in MSAL.js 2.x service wrapper & helper for SPAs
//   v2.1.0 - Ben Coleman 2019
//   Updated 2021 - Switched to @azure/msal-browser
// ----------------------------------------------------------------------------

import * as msal from '@azure/msal-browser'

// MSAL object used for signing in users with MS identity platform
let msalApp

export default {
  //
  // Configure with clientId or empty string/null to set in "demo" mode
  //
  async configure(clientId, tenantId = 'common') {
    // Can only call configure once
    if (msalApp) {
      return
    }

    // Can't configure if clientId blank/null/undefined
    if (!clientId) {
      return
    }

    const config = {
      auth: {
        clientId: clientId,
        redirectUri: window.location.origin,
        authority: 'https://login.microsoftonline.com/' + tenantId
      },
      cache: {
        cacheLocation: 'localStorage'
      }
      // Only uncomment when you *really* need to debug what is going on in MSAL
      /* system: {
        logger: new msal.Logger(
          (logLevel, msg) => { console.log(msg) },
          {
            level: msal.LogLevel.Verbose
          }
        )
      } */
    }
    console.log('### Azure AD sign-in: enabled\n', config)

    // Create our shared/static MSAL app object
    msalApp = new msal.PublicClientApplication(config)
  },

  async login(scopes = ['User.Read', 'Mail.ReadBasic', 'Mail.Read']) {
    if (!msalApp) {
      return
    }

    await msalApp.loginPopup({
      scopes,
      prompt: 'select_account'
    })
  },

  user() {
    if (!msalApp) {
      return null
    }

    const currentAccounts = msalApp.getAllAccounts()
    if (!currentAccounts || currentAccounts.length === 0) {
      // No user signed in
      return null
    } else if (currentAccounts.length > 1) {
      return currentAccounts[0]
    } else {
      return currentAccounts[0]
    }
  },

}