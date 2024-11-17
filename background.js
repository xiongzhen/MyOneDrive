chrome.action.onClicked.addListener((tab) => {
    chrome.tabs.create({
        url: "my-onedrive.html",
    }, (tab) => {
        let get_access_token = (doProceed) => {
            let auth_url = 'https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize?';

            // https://dkjkjcfpdeehldedmboaodphlnlciemh.chromiumapp.org/OAuth2
            let redirect_url = chrome.identity.getRedirectURL('OAuth2');

            let auth_param = {
                client_id: "ff697d5e-37f8-41dc-9956-6c62b89c2ff5",
                response_type: "token",
                redirect_url: redirect_url,
                response_mode: "fragment",
                scope: "User.Read Files.ReadWrite.All",
                state: "314159265"
            };
            const url = new URLSearchParams(Object.entries(auth_param));
            url.toString();
            auth_url += url;
            chrome.identity.launchWebAuthFlow({
                    url: auth_url,
                    interactive: true
                }, (response_url) => {
                    let queryString = response_url.split('#')[1];
                    const urlParams = new URLSearchParams(queryString);
                    const access_token = urlParams.get('access_token');
                    let storageCache = {
                        access_token
                    };
                    chrome.storage.local.set({access_token: access_token}, function() {
                        //if(doProceed) {
                        //    chrome.scripting.executeScript({
                        //        target: {tabId: tab.id},
                        //        files: ['content.js']
                        //    });
                        //}
                    });
                }
            );
        };

        get_access_token(true);
        setInterval(get_access_token, 3300 * 1000, false);   // 3600 seonds - 5 minutes = 3300 seconds
    });
});