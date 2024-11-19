browser.tabs.create({
    url: "my-onedrive.html",
}, (tab) => {
    let get_access_token = (refresh) => {
        let auth_url = 'https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize?';

        // https://dkjkjcfpdeehldedmboaodphlnlciemh.chromiumapp.org/OAuth2
        let redirect_url = browser.identity.getRedirectURL('OAuth2');

        let auth_param = {
            client_id: "15907424-9c63-4fc4-bc77-8e2976496a6c",
            response_type: "token",
            redirect_url: redirect_url,
            response_mode: "fragment",
            scope: "User.Read Files.ReadWrite.All",
            state: "314159265"
        };
        const url = new URLSearchParams(Object.entries(auth_param));
        url.toString();
        auth_url += url;
        try {
            browser.identity.launchWebAuthFlow({
                    url: auth_url,
                    interactive: true
                }, (response_url) => {
                    let queryString = response_url.split('#')[1];
                    const urlParams = new URLSearchParams(queryString);
                    const access_token = urlParams.get('access_token');
                    let storageCache = {
                        access_token
                    };
                    browser.storage.local.set({access_token: access_token}, function() {
                    });
                }
            );
        } catch(err) {
            get_access_token();
        }
    };

    get_access_token(true);
    setInterval(get_access_token, 3300 * 1000, false);   // 3600 seonds - 5 minutes = 3300 seconds
});