/* =========================================================
   CCM — STEP 11C
   GOOGLE WORKSPACE AUTHENTICATION
========================================================= */


/* =========================================================
   GOOGLE OAUTH CONFIGURATION
========================================================= */

const GOOGLE_CLIENT_ID =
    "644454810051-pa1247vs2636vb5o0fab1loin87tu5vd.apps.googleusercontent.com";


const CCM_SCRIPT_ID =
    "1U-kT27QuT7h98HNKHlseXJubYRIknWkTwtp1MYx5n60TYcvJYKJR8SAK";


const CCM_OAUTH_SCOPES = [
    "https://www.googleapis.com/auth/script.scriptapp",
    "https://www.googleapis.com/auth/script.send_mail",
    "https://www.googleapis.com/auth/script.external_request",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/drive.readonly"
].join(" ");


/* =========================================================
   CCM AUTH STATE
========================================================= */

let CCM_ID_TOKEN = null;
let CCM_ACCESS_TOKEN = null;
let CCM_TOKEN_CLIENT = null;


/* =========================================================
   START
========================================================= */

document.addEventListener(
    "DOMContentLoaded",
    () => {

        initializeCCMOAuth();
        initializeGoogleLogin();

    }
);


/* =========================================================
   INITIALIZE CCM OAUTH
========================================================= */

function initializeCCMOAuth() {

    if (
        !window.google ||
        !google.accounts ||
        !google.accounts.oauth2
    ) {

        return false;

    }


    try {

        CCM_TOKEN_CLIENT =
            google.accounts.oauth2.initTokenClient({

                client_id:
                    GOOGLE_CLIENT_ID,

                scope:
                    CCM_OAUTH_SCOPES,

                callback:
                    handleCCMAccessToken

            });


        console.log(
            "CCM: OAuth token client initialized."
        );


        return true;

    } catch (error) {

        console.error(
            "CCM OAuth initialization failed:",
            error
        );


        showLoginError(
            "Unable to initialize CCM authorization."
        );


        return false;

    }

}


/* =========================================================
   INITIALIZE GOOGLE LOGIN
========================================================= */

function initializeGoogleLogin() {

    const loginContainer =
        document.getElementById(
            "googleLogin"
        );


    if (!loginContainer) {

        console.error(
            "CCM: Google login container not found."
        );

        return;

    }


    waitForGoogleIdentityServices(
        loginContainer
    );

}


/* =========================================================
   WAIT FOR GOOGLE GIS
========================================================= */

function waitForGoogleIdentityServices(
    loginContainer,
    attempts = 0
) {

    if (
        window.google &&
        google.accounts &&
        google.accounts.id &&
        google.accounts.oauth2
    ) {

        if (!CCM_TOKEN_CLIENT) {

            initializeCCMOAuth();

        }


        renderGoogleButton(
            loginContainer
        );


        return;

    }


    if (attempts >= 50) {

        showLoginError(
            "Google authentication service could not be loaded."
        );


        return;

    }


    setTimeout(
        () => {

            waitForGoogleIdentityServices(
                loginContainer,
                attempts + 1
            );

        },
        200
    );

}


/* =========================================================
   RENDER GOOGLE SIGN-IN BUTTON
========================================================= */

function renderGoogleButton(
    loginContainer
) {

    try {

        google.accounts.id.initialize({

            client_id:
                GOOGLE_CLIENT_ID,

            callback:
                handleGoogleCredential,

            auto_select:
                false,

            cancel_on_tap_outside:
                true

        });


        google.accounts.id.renderButton(

            loginContainer,

            {

                type:
                    "standard",

                theme:
                    "outline",

                size:
                    "large",

                text:
                    "signin_with",

                shape:
                    "rectangular",

                logo_alignment:
                    "left",

                width:
                    320

            }

        );


        console.log(
            "CCM: Google Sign-In button rendered."
        );


    } catch (error) {

        console.error(
            "CCM Google Login Error:",
            error
        );


        showLoginError(
            "Unable to initialize Google authentication."
        );

    }

}


/* =========================================================
   GOOGLE ID TOKEN
========================================================= */

function handleGoogleCredential(
    response
) {

    console.log(
        "CCM: Google authentication successful."
    );


    if (
        !response ||
        !response.credential
    ) {

        showLoginError(
            "Google did not return an authentication credential."
        );


        return;

    }


    /*
       Store ID token only in memory.
    */

    CCM_ID_TOKEN =
        response.credential;


    /*
       Request OAuth access token.

       ID token is NOT used for scripts.run.
    */

    requestCCMAccessToken();

}


/* =========================================================
   REQUEST CCM ACCESS TOKEN
========================================================= */

function requestCCMAccessToken() {

    if (!CCM_TOKEN_CLIENT) {

        if (!initializeCCMOAuth()) {

            showLoginError(
                "CCM authorization could not be initialized."
            );


            return;

        }

    }


    try {

        CCM_TOKEN_CLIENT.requestAccessToken({

            prompt:
                ""

        });


    } catch (error) {

        console.error(
            "CCM access token request failed:",
            error
        );


        showLoginError(
            "Unable to obtain Google Workspace authorization."
        );

    }

}


/* =========================================================
   ACCESS TOKEN CALLBACK
========================================================= */

function handleCCMAccessToken(
    response
) {

    if (
        !response ||
        response.error
    ) {

        console.error(
            "CCM OAuth error:",
            response
        );


        showLoginError(
            "Google authorization was denied or failed."
        );


        return;

    }


    if (!response.access_token) {

        showLoginError(
            "Google did not return an access token."
        );


        return;

    }


    /*
       Store access token only in memory.
    */

    CCM_ACCESS_TOKEN =
        response.access_token;


    console.log(
        "CCM: Google OAuth access token received."
    );


    authenticateCCMBackend();

}


/* =========================================================
   CCM API EXECUTOR
========================================================= */

async function ccmExecute(
    action,
    params = {}
) {

    if (!CCM_ACCESS_TOKEN) {

        throw new Error(
            "CCM access token is not available."
        );

    }


    const response =
        await fetch(
            "https://script.googleapis.com/v1/scripts/" +
            encodeURIComponent(CCM_SCRIPT_ID) +
            ":run",
            {

                method:
                    "POST",

                headers:
                    {

                        "Authorization":
                            "Bearer " +
                            CCM_ACCESS_TOKEN,

                        "Content-Type":
                            "application/json"

                    },

                body:
                    JSON.stringify({

                        function:
                            "ccmExecute",

                        parameters:
                            [
                                action,
                                params
                            ]

                    })

            }
        );


    const data =
        await response.json();


    console.log(
        "CCM API [" + action + "]:",
        data
    );


    if (!response.ok) {

        throw new Error(
            data.error?.message ||
            "CCM API request failed."
        );

    }


    if (data.error) {

        throw new Error(
            data.error.message ||
            "CCM Apps Script execution failed."
        );

    }


    if (
        !data.response ||
        !data.response.result
    ) {

        throw new Error(
            "CCM returned an invalid API response."
        );

    }


    const result =
        data.response.result;


    if (!result.success) {

        throw new Error(
            result.error ||
            "CCM request was rejected."
        );

    }


    return result;

}


/* =========================================================
   CALL CCM BACKEND
========================================================= */

async function authenticateCCMBackend() {

    if (!CCM_ACCESS_TOKEN) {

        showLoginError(
            "CCM access token is missing."
        );


        return;

    }


    try {

        showLoginConnecting();


        const result =
            await ccmExecute(
                "me",
                {}
            );


        console.log(
            "CCM backend response:",
            result
        );


        if (!result.success) {

            showLoginError(
                result.error ||
                "You are not authorized to access CCM."
            );


            return;

        }


        if (
            result.authenticated !== true
        ) {

            showLoginError(
                "Google authentication could not be verified by CCM."
            );


            return;

        }


        if (
            result.authorized !== true
        ) {

            showLoginError(
                "You do not have permission to access CCM."
            );


            return;

        }


        /*
           Store authenticated user information.
        */

        window.CCM_CURRENT_USER =
            result.user || null;


        window.CCM_PERMISSIONS =
            result.permissions || [];


        console.log(
            "CCM authentication successful:",
            result
        );


        showLoginSuccess(
            result.user
        );


    } catch (error) {

        console.error(
            "CCM backend connection error:",
            error
        );


        showLoginError(
            error.message ||
            "Unable to connect to the CCM backend."
        );

    }

}


/* =========================================================
   LOGIN CONNECTING
========================================================= */

function showLoginConnecting() {

    const status =
        document.getElementById(
            "loginStatus"
        );


    if (status) {

        status.className =
            "login-status success";


        status.textContent =
            "Google authenticated. Verifying CCM access...";

    }

}


/* =========================================================
   LOGIN SUCCESS
========================================================= */

function showLoginSuccess(
    user
) {

    const status =
        document.getElementById(
            "loginStatus"
        );


    if (status) {

        status.className =
            "login-status success";


        status.textContent =
            "CCM authentication successful. Loading dashboard...";

    }


    /*
       Populate user information.
    */

    if (user) {

        const userName =
            document.getElementById(
                "userName"
            );


        const userRole =
            document.getElementById(
                "userRole"
            );


        if (userName) {

            userName.textContent =
                user.Name ||
                user.Email ||
                "User";

        }


        if (userRole) {

            userRole.textContent =
                user.Role_ID ||
                "";

        }

    }


    setTimeout(
        () => {

            openCCMApplication();

        },
        500
    );

}


/* =========================================================
   OPEN CCM
========================================================= */

function openCCMApplication() {

    const loginScreen =
        document.getElementById(
            "loginScreen"
        );


    if (loginScreen) {

        loginScreen.style.display =
            "none";

    }


    initNavigation();
    initMobileMenu();

}


/* =========================================================
   LOGIN ERROR
========================================================= */

function showLoginError(
    message
) {

    const status =
        document.getElementById(
            "loginStatus"
        );


    if (!status) {

        return;

    }


    status.className =
        "login-status error";


    status.textContent =
        message;

}


/* =========================================================
   NAVIGATION
========================================================= */

function initNavigation() {

    const navItems =
        document.querySelectorAll(
            ".nav-item"
        );


    const pages =
        document.querySelectorAll(
            ".page"
        );


    const pageTitle =
        document.getElementById(
            "pageTitle"
        );


    navItems.forEach(
        item => {

            item.addEventListener(
                "click",
                () => {

                    const page =
                        item.dataset.page;


                    navItems.forEach(
                        nav => {

                            nav.classList.remove(
                                "active"
                            );

                        }
                    );


                    item.classList.add(
                        "active"
                    );


                    pages.forEach(
                        section => {

                            section.classList.remove(
                                "active-page"
                            );

                        }
                    );


                    const target =
                        document.getElementById(
                            page + "Page"
                        );


                    if (target) {

                        target.classList.add(
                            "active-page"
                        );

                    }


                    const title =
                        item.querySelector(
                            "span:last-child"
                        );


                    if (
                        title &&
                        pageTitle
                    ) {

                        pageTitle.textContent =
                            title.textContent;

                    }


                    const sidebar =
                        document.getElementById(
                            "sidebar"
                        );


                    if (sidebar) {

                        sidebar.classList.remove(
                            "open"
                        );

                    }

                }
            );

        }
    );

}


/* =========================================================
   MOBILE MENU
========================================================= */

function initMobileMenu() {

    const menuButton =
        document.getElementById(
            "menuButton"
        );


    const sidebar =
        document.getElementById(
            "sidebar"
        );


    if (
        !menuButton ||
        !sidebar
    ) {

        return;

    }


    menuButton.addEventListener(
        "click",
        () => {

            sidebar.classList.toggle(
                "open"
            );

        }
    );

}
