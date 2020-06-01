 // Config object to be passed to Msal on creation.
 // For a full list of msal.js configuration parameters, 
 // visit https://azuread.github.io/microsoft-authentication-library-for-js/docs/msal/modules/_authenticationparameters_.html
 const msalConfig = {
     auth: {
         clientId: "7c60adbf-59ce-495e-ad97-ec3b66135f24",
         authority: "https://login.microsoftonline.com/8900c5e2-270f-4ca0-9a2c-85889720d198",
         redirectUri: "http://localhost:3000/",
     },
     cache: {
         cacheLocation: "sessionStorage", // This configures where your cache will be stored
         storeAuthStateInCookie: false, // Set this to "true" if you are having issues on IE11 or Edge
     }
 };

 // Add here scopes for id token to be used at MS Identity Platform endpoints.
 const loginRequest = {
     scopes: ["openid", "profile", "User.Read"]
 };

 // Add here scopes for access token to be used at MS Graph API endpoints.
 const tokenRequest = {
     scopes: ["Sites.ReadWrite.All", "Sites.Read.All", "Sites.Manage.All"]
 };


 // Helper function to call MS Graph API endpoint 
 // using authorization bearer token scheme

 // Add here the endpoints for MS Graph API services you would like to use.
 const graphConfig = {
     graphMeEndpoint: "https://graph.microsoft.com/v1.0/me",
     graphSPpoint: "https://graph.microsoft.com/v1.0/sites/wajesmartltd.sharepoint.com,0f07b3e4-adbe-4c10-8b5e-809290795fb3,a5e17ab5-3521-45ae-af3c-2eb9eb946063/lists/17a21815-3e31-41c0-928b-3f76e6ea8748/items"
 };



 function callMSGraphPost(endpoint, token, data, callback) {

     fetch(endpoint, {
         method: "POST",
         headers: {
             'Authorization': 'Bearer ' + token,
             'Content-type': 'application/json'
         },
         body: JSON.stringify(data)
     }).then(data => {
         return data.json()
     }).then(res => {
         console.log(res);
     }).catch(err => {
         console.log(err)
     })
 }

 function displayList(data, endpoint) {
     console.log(data);
 }

 // Create the main myMSALObj instance
 // configuration parameters are located at authConfig.js
 const myMSALObj = new Msal.UserAgentApplication(msalConfig);

 function signIn() {
     myMSALObj.loginPopup(loginRequest)
         .then(loginResponse => {
             // console.log('id_token acquired at: ' + new Date().toString());
             // console.log(loginResponse);

             if (myMSALObj.getAccount()) {
                 // showWelcomeMessage(myMSALObj.getAccount());
             }
         }).catch(error => {
             console.log(error);
         });
 }

 function signOut() {
     myMSALObj.logout();
 }

 function getTokenPopup(request) {
     return myMSALObj.acquireTokenSilent(request)
         .catch(error => {
             console.log(error);
             console.log("silent token acquisition fails. acquiring token using popup");

             // fallback to interaction when silent call fails
             return myMSALObj.acquireTokenPopup(request)
                 .then(tokenResponse => {
                     return tokenResponse;
                 }).catch(error => {
                     console.log(error);
                 });
         });
 }


 document.getElementById("form").addEventListener("submit", handleList);

 function handleList(submitEvent) {

     submitEvent.preventDefault();
     signIn()


     if (myMSALObj.getAccount()) {
         getTokenPopup(tokenRequest)
             .then(response => {
                 const data = {
                     fields: {
                         Request: document.getElementById('request').value,
                         ProjectPurpose: document.getElementById('purpose').value,
                         NameofVendor: document.getElementById('vendorname').value,
                         InitiatorsAccountName: document.getElementById('accountname').value,
                         InitiatorsAccountNumber: document.getElementById('accountnumber').value,
                         //  VendorBankName: document.getElementById('bankname').value,
                         PaymentTerm: document.getElementById('paymentterm').value,
                         //  AdvancePayment: document.getElementById('advance').value,
                         //  Total_Invoice_Amount: document.getElementById('invoice').value,
                         DateofRequest: document.getElementById('date').value,
                         Authorizer: document.getElementById('authorizer').value,
                         Attachments: document.getElementById('file').value,

                     }
                 };
                 console.log(data)
                     // callMSGraph(graphConfig.graphSPpoint, response.accessToken, displayList);
                 callMSGraphPost(graphConfig.graphSPpoint, response.accessToken, data, displayList)

             }).catch(error => {
                 console.log(error);
             });
     }
 }