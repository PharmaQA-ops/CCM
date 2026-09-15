/* =========================================================
   CCM — STEP 11C
   GOOGLE WORKSPACE AUTHENTICATION
========================================================= */


/*
   IMPORTANT:
   Replace this with your Google OAuth 2.0
   Web Application Client ID.
*/

const GOOGLE_CLIENT_ID =
    "644454810051-pa1247vs2636vb5o0fab1loin87tu5vd.apps.googleusercontent.com";


let CCM_ID_TOKEN = null;


/* =========================================================
   START
========================================================= */

document.addEventListener(
    "DOMContentLoaded",
    () => {

        initializeGoogleLogin();

    }
);


/* =========================================================
   GOOGLE LOGIN
========================================================= */

function initializeGoogleLogin() {

    const loginContainer =
        document.getElementById("googleLogin");

    if (!loginContainer) {

        console.error(
            "CCM: Google login container not found."
        );

        return;
    }


    /*
       Google Identity Services may load
       slightly after DOMContentLoaded.
    */

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
        google.accounts.id
    ) {

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
   RENDER GOOGLE BUTTON
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

                type: "standard",

                theme: "outline",

                size: "large",

                text: "signin_with",

                shape: "rectangular",

                logo_alignment: "left",

                width: 320

            }

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
   GOOGLE CREDENTIAL
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
       Keep the token in memory only.

       We will pass it to the CCM backend
       in the next authentication step.

       We are deliberately NOT putting
       the token into localStorage.
    */

    CCM_ID_TOKEN =
        response.credential;


    showLoginSuccess();


    /*
       Do NOT call the CCM API yet.

       11D will verify/send the identity
       to the CCM backend.
    */

}


/* =========================================================
   LOGIN SUCCESS
========================================================= */

function showLoginSuccess() {

    const status =
        document.getElementById(
            "loginStatus"
        );

    if (status) {

        status.className =
            "login-status success";

        status.textContent =
            "Google authentication successful. Connecting to CCM...";

    }


    /*
       For now, after Google authentication,
       reveal the CCM application.

       Backend authorization will be added
       in 11D.
    */

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


    /*
       Initialize the dashboard shell
       only after successful login.
    */

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
