// ----------------------------------------------------------------------------
// Copyright (c) Ben Coleman, 2020
// Licensed under the MIT License.
//
// Set of methods to call the beta Microsoft Graph API, using REST and fetch
// Requires auth.js
// ----------------------------------------------------------------------------

import auth from './auth'

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0'
const GRAPH_SCOPES =  ['User.Read', 'Mail.ReadBasic', 'Mail.Read']

let accessToken

export default {

    async getEmails() {
        let response = await callGraph('/me/messages')
        if (response) {
            let data = await response.json()
            return data.value
        }
    },

}

//
// Common fetch wrapper (private)
//
async function callGraph(apiPath) {
    // Acquire an access token to call APIs (like Graph)
    // Safe to call repeatedly as MSAL caches tokens locally
    accessToken = await auth.acquireToken(GRAPH_SCOPES)

    let response = await fetch(`${GRAPH_BASE}${apiPath}`, {
        headers: { authorization: `bearer ${accessToken}` }
    })

    if (!response.ok) {
        throw new Error(`Call to ${GRAPH_BASE}${apiPath} failed: ${response.statusText}`)
    }

    return response
}