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
   LOAD LIVE DASHBOARD
========================================================= */

async function loadDashboard() {

    try {

        console.log(
            "CCM: Loading live dashboard..."
        );

        const data =
            await ccmExecute(
                "dashboard",
                {}
            );


        if (!data || !data.success) {

            throw new Error(
                data?.error ||
                "Dashboard data could not be loaded."
            );

        }


        console.log(
            "CCM: Dashboard data loaded:",
            data
        );

/* =====================================================
           KPI DATA
        ===================================================== */

        const kpi =
            data.kpi || {};


        setElementText(
            "totalCertifications",
            kpi.total ?? 0
        );


        setElementText(
            "activeCertifications",
            kpi.active ?? 0
        );


        setElementText(
            "renewalCertifications",
            kpi.renewal ?? 0
        );


        setElementText(
            "expiredCertifications",
            kpi.expired ?? 0
        );


        setElementText(
            "mandatoryCertifications",
            kpi.mandatory ?? 0
        );


        setElementText(
            "complianceRate",
            (kpi.compliance ?? 0) + "%"
        );


        /* =====================================================
           EXPIRY RISK
        ===================================================== */

        const risk =
            data.expiryRisk || {};


        updateRiskRow(
            "risk07",
            "risk07Value",
            risk["0-7"] || 0
        );


        updateRiskRow(
            "risk815",
            "risk815Value",
            risk["8-15"] || 0
        );


        updateRiskRow(
            "risk1630",
            "risk1630Value",
            risk["16-30"] || 0
        );


        updateRiskRow(
            "risk3160",
            "risk3160Value",
            risk["31-60"] || 0
        );


        updateRiskRow(
            "risk6190",
            "risk6190Value",
            risk["61-90"] || 0
        );


        updateRiskRow(
            "risk90",
            "risk90Value",
            risk["90+"] || 0
        );


        updateRiskRow(
            "riskExpired",
            "riskExpiredValue",
            risk["EXPIRED"] || 0
        );


        console.log(
            "CCM: Live dashboard rendered successfully."
        );


    } catch (error) {

        console.error(
            "CCM dashboard loading failed:",
            error
        );

    }

}


/* =========================================================
   SAFE TEXT UPDATE
========================================================= */

function setElementText(
    elementId,
    value
) {

    const element =
        document.getElementById(
            elementId
        );


    if (element) {

        element.textContent =
            value;

    }

}


/* =========================================================
   UPDATE EXPIRY RISK ROW
========================================================= */

function updateRiskRow(
    barId,
    valueId,
    value
) {

    const bar =
        document.getElementById(
            barId
        );


    const valueElement =
        document.getElementById(
            valueId
        );


    if (valueElement) {

        valueElement.textContent =
            value;

    }


    if (bar) {

        /*
           Width is relative to the largest
           expiry-risk category.
        */

        const allRiskValues = [

            getRiskValue("risk07Value"),
            getRiskValue("risk815Value"),
            getRiskValue("risk1630Value"),
            getRiskValue("risk3160Value"),
            getRiskValue("risk6190Value"),
            getRiskValue("risk90Value"),
            getRiskValue("riskExpiredValue")

        ];


        const maxValue =
            Math.max(
                ...allRiskValues,
                1
            );


        const percentage =
            Math.round(
                (value / maxValue) * 100
            );


        bar.style.width =
            percentage + "%";

    }

}


/* =========================================================
   GET RISK VALUE
========================================================= */

function getRiskValue(
    elementId
) {

    const element =
        document.getElementById(
            elementId
        );


    if (!element) {

        return 0;

    }


    const value =
        parseInt(
            element.textContent,
            10
        );


    return Number.isFinite(value)
        ? value
        : 0;

}

/* =========================================================
   CCM — STEP 11E
   CERTIFICATION OVERVIEW
========================================================= */

let CCM_CERTIFICATES = [];


/* =========================================================
   LOAD CERTIFICATIONS
========================================================= */

async function loadCertificates() {

    const tableBody =
        document.getElementById(
            "certificateTableBody"
        );


    if (tableBody) {

        tableBody.innerHTML = `
            <tr>
                <td colspan="9" class="table-loading">
                    Loading certification records...
                </td>
            </tr>
        `;

    }


    try {

        const result =
            await ccmExecute(
                "certificates",
                {}
            );


        if (
            !result ||
            !result.success
        ) {

            throw new Error(
                result?.error ||
                "Certification records could not be loaded."
            );

        }


        CCM_CERTIFICATES =
            Array.isArray(result.data)
                ? result.data
                : [];


        console.log(
            "CCM: Certification records loaded:",
            CCM_CERTIFICATES
        );


        populateCertificateDepartments();

        renderCertificates();


    } catch (error) {

        console.error(
            "CCM certification loading failed:",
            error
        );


        if (tableBody) {

            tableBody.innerHTML = `
                <tr>
                    <td colspan="9" class="table-error">
                        Unable to load certification records.
                    </td>
                </tr>
            `;

        }

    }

}


/* =========================================================
   POPULATE DEPARTMENT FILTER
========================================================= */

function populateCertificateDepartments() {

    const select =
        document.getElementById(
            "certificateDepartmentFilter"
        );


    if (!select) {

        return;

    }


    const departments =
        [
            ...new Set(
                CCM_CERTIFICATES
                    .map(
                        certificate =>
                            String(
                                certificate.Department || ""
                            ).trim()
                    )
                    .filter(Boolean)
            )
        ]
        .sort(
            (a, b) =>
                a.localeCompare(b)
        );


    select.innerHTML = `
        <option value="">
            All Departments
        </option>
    `;


    departments.forEach(
        department => {

            const option =
                document.createElement(
                    "option"
                );


            option.value =
                department;


            option.textContent =
                department;


            select.appendChild(
                option
            );

        }
    );

}


/* =========================================================
   RENDER CERTIFICATIONS
========================================================= */

function renderCertificates() {

    const tableBody =
        document.getElementById(
            "certificateTableBody"
        );


    const summary =
        document.getElementById(
            "certificateResultSummary"
        );


    if (!tableBody) {

        return;

    }


    const search =
        String(
            document.getElementById(
                "certificateSearch"
            )?.value || ""
        )
        .trim()
        .toLowerCase();


    const status =
        String(
            document.getElementById(
                "certificateStatusFilter"
            )?.value || ""
        )
        .trim()
        .toUpperCase();


    const department =
        String(
            document.getElementById(
                "certificateDepartmentFilter"
            )?.value || ""
        )
        .trim()
        .toLowerCase();


    const filtered =
        CCM_CERTIFICATES.filter(
            certificate => {

                const searchText = [

                    certificate.Certificate_ID,
                    certificate.Employee_ID,
                    certificate.Employee_Name,
                    certificate.Certification_Name,
                    certificate.Department,
                    certificate.Issuing_Body,
                    certificate.Certificate_Number

                ]
                .join(" ")
                .toLowerCase();


                const matchesSearch =
                    !search ||
                    searchText.includes(
                        search
                    );


                const matchesStatus =
                    !status ||
                    String(
                        certificate.Status || ""
                    )
                    .toUpperCase() ===
                    status;


                const matchesDepartment =
                    !department ||
                    String(
                        certificate.Department || ""
                    )
                    .toLowerCase() ===
                    department;


                return (
                    matchesSearch &&
                    matchesStatus &&
                    matchesDepartment
                );

            }
        );


    if (summary) {

        summary.textContent =
            `${filtered.length} of ${CCM_CERTIFICATES.length} certification records`;

    }


    if (!filtered.length) {

        tableBody.innerHTML = `
            <tr>
                <td
                    colspan="9"
                    class="table-empty"
                >
                    No certification records found.
                </td>
            </tr>
        `;

        return;

    }


    tableBody.innerHTML =
        filtered
            .map(
                certificate =>
                    createCertificateRow(
                        certificate
                    )
            )
            .join("");


    /*
       Attach action handlers after rendering.
    */

    tableBody
        .querySelectorAll(
            "[data-certificate-id]"
        )
        .forEach(
            button => {

                button.addEventListener(
                    "click",
                    () => {

                        const certificateId =
                            button.dataset.certificateId;


                        openCertificate360(
                            certificateId
                        );

                    }
                );

            }
        );

}


/* =========================================================
   CREATE CERTIFICATE ROW
========================================================= */

function createCertificateRow(
    certificate
) {

    const certificateId =
        escapeHtml(
            certificate.Certificate_ID
        );


    const employee =
        escapeHtml(
            certificate.Employee_Name ||
            certificate.Employee_ID ||
            "-"
        );


    const certification =
        escapeHtml(
            certificate.Certification_Name ||
            "-"
        );


    const department =
        escapeHtml(
            certificate.Department ||
            "-"
        );


    const issueDate =
        escapeHtml(
            certificate.Issue_Date ||
            "-"
        );


    const expiryDate =
        escapeHtml(
            certificate.Expiry_Date ||
            "-"
        );


    const status =
        String(
            certificate.Status || ""
        )
        .toUpperCase();


    const renewal =
        String(
            certificate.Renewal_Status || ""
        )
        .toUpperCase();


    return `
        <tr>

            <td>
                <span class="certificate-id">
                    ${certificateId}
                </span>
            </td>


            <td>
                <strong>
                    ${employee}
                </strong>
            </td>


            <td>
                ${certification}
            </td>


            <td>
                ${department}
            </td>


            <td>
                ${issueDate}
            </td>


            <td>
                ${expiryDate}
            </td>


            <td>
                <span class="
                    certificate-status
                    ${getStatusClass(status)}
                ">
                    ${escapeHtml(
                        status || "-"
                    )}
                </span>
            </td>


            <td>
                <span class="
                    certificate-renewal
                    ${getRenewalClass(renewal)}
                ">
                    ${escapeHtml(
                        renewal || "-"
                    )}
                </span>
            </td>


            <td>

                <button
                    type="button"
                    class="table-action"
                    data-certificate-id="${certificateId}"
                >
                    View
                </button>

            </td>

        </tr>
    `;

}


/* =========================================================
   STATUS CLASS
========================================================= */

function getStatusClass(
    status
) {

    switch (status) {

        case "ACTIVE":
            return "status-active";

        case "EXPIRED":
            return "status-expired";

        case "INACTIVE":
            return "status-inactive";

        default:
            return "status-neutral";

    }

}


/* =========================================================
   RENEWAL CLASS
========================================================= */

function getRenewalClass(
    renewal
) {

    if (
        renewal.includes("EXPIRED")
    ) {

        return "renewal-expired";

    }


    if (
        renewal.includes("URGENT") ||
        renewal.includes("DUE") ||
        renewal.includes("7") ||
        renewal.includes("15") ||
        renewal.includes("30")
    ) {

        return "renewal-warning";

    }


    if (
        renewal.includes("RENEWED")
    ) {

        return "renewal-completed";

    }


    return "renewal-neutral";

}


/* =========================================================
   HTML ESCAPE
========================================================= */

function escapeHtml(
    value
) {

    return String(
        value ?? ""
    )
    .replace(
        /&/g,
        "&amp;"
    )
    .replace(
        /</g,
        "&lt;"
    )
    .replace(
        />/g,
        "&gt;"
    )
    .replace(
        /"/g,
        "&quot;"
    )
    .replace(
        /'/g,
        "&#039;"
    );

}


/* =========================================================
   CERTIFICATE 360
========================================================= */

async function openCertificate360(
    certificateId
) {

    console.log(
        "CCM: Opening Certificate 360:",
        certificateId
    );


    /*
       11E only opens the record.
       Full Certificate 360 UI will be built
       in the next certification-management step.
    */

    try {

        const result =
            await ccmExecute(
                "certificate360",
                {
                    certificateId:
                        certificateId
                }
            );


        if (
            !result ||
            !result.success
        ) {

            throw new Error(
                result?.error ||
                "Certificate details could not be loaded."
            );

        }


        console.log(
            "CCM: Certificate 360:",
            result
        );


        /*
           Temporary production-safe behaviour:
           show the record in a controlled dialog.
        */

        showCertificateDetails(
            result
        );


    } catch (error) {

        console.error(
            "CCM Certificate 360 failed:",
            error
        );

        alert(
            error.message ||
            "Unable to open certificate."
        );

    }

}


/* =========================================================
   CERTIFICATE DETAILS
========================================================= */

function showCertificateDetails(
    result
) {

    const certificate =
        result?.certificate || {};

    const employee =
        result?.employee || {};

    const certificateDocument =
        result?.document || {};

    /*
       Remove any existing Certificate 360 modal.
    */

    const existingModal =
        document.getElementById(
            "certificate360Modal"
        );

    if (existingModal) {
        existingModal.remove();
    }


    /*
       Safe display helper.
    */

    const safe = value =>
        escapeHtml(
            value === null ||
            value === undefined ||
            value === ""
                ? "-"
                : value
        );


    const employeeName =
        employee.Employee_Name ||
        certificate.Employee_Name ||
        "-";


    const employeeId =
        employee.Employee_ID ||
        certificate.Employee_ID ||
        "-";


    const status =
        String(
            certificate.Status || "-"
        ).toUpperCase();


    const renewalStatus =
        String(
            certificate.Renewal_Status || "-"
        ).toUpperCase();


    const documentId =
        certificateDocument.Document_ID ||
        "-";


    const classification =
        certificateDocument.Classification ||
        "-";


    const documentName =
        certificateDocument.Document_Name ||
        certificateDocument.File_Name ||
        "-";


    const downloadAllowed =
        certificateDocument.Download_Allowed === true;


    /*
       Create Certificate 360 modal.
    */

    const modal =
        document.createElement("div");

    modal.id =
        "certificate360Modal";

    modal.className =
        "certificate360-modal";


    modal.innerHTML = `

        <div
            class="certificate360-backdrop"
            data-certificate360-close
        ></div>


        <section
            class="certificate360-dialog"
            role="dialog"
            aria-modal="true"
            aria-labelledby="certificate360Title"
        >

            <header class="certificate360-header">

                <div>

                    <span class="certificate360-eyebrow">
                        CERTIFICATION RECORD
                    </span>

                    <h2 id="certificate360Title">
                        Certificate 360
                    </h2>

                    <p>
                        Complete certification record
                    </p>

                </div>


                <button
                    type="button"
                    class="certificate360-close"
                    id="certificate360Close"
                    aria-label="Close"
                >
                    ×
                </button>

            </header>


            <div class="certificate360-body">

                <div class="certificate360-summary">

                    <div class="certificate360-summary-icon">
                        ✓
                    </div>

                    <div>

                        <span>
                            Certificate ID
                        </span>

                        <strong>
                            ${safe(
                                certificate.Certificate_ID
                            )}
                        </strong>

                    </div>

                </div>


                <section class="certificate360-section">

                    <div class="certificate360-section-title">
                        Certificate Details
                    </div>

                    <div class="certificate360-grid">

                        <div class="certificate360-field">
                            <span>Certification</span>
                            <strong>
                                ${safe(
                                    certificate.Certification_Name
                                )}
                            </strong>
                        </div>

                        <div class="certificate360-field">
                            <span>Certificate Number</span>
                            <strong>
                                ${safe(
                                    certificate.Certificate_Number
                                )}
                            </strong>
                        </div>

                        <div class="certificate360-field">
                            <span>Issuing Body</span>
                            <strong>
                                ${safe(
                                    certificate.Issuing_Body
                                )}
                            </strong>
                        </div>

                        <div class="certificate360-field">
                            <span>Department</span>
                            <strong>
                                ${safe(
                                    certificate.Department
                                )}
                            </strong>
                        </div>

                        <div class="certificate360-field">
                            <span>Issue Date</span>
                            <strong>
                                ${safe(
                                    certificate.Issue_Date
                                )}
                            </strong>
                        </div>

                        <div class="certificate360-field">
                            <span>Expiry Date</span>
                            <strong>
                                ${safe(
                                    certificate.Expiry_Date
                                )}
                            </strong>
                        </div>

                        <div class="certificate360-field">
                            <span>Status</span>
                            <strong
                                class="certificate360-value-badge ${getStatusClass(status)}"
                            >
                                ${safe(status)}
                            </strong>
                        </div>

                        <div class="certificate360-field">
                            <span>Renewal Status</span>
                            <strong
                                class="certificate360-value-badge ${getRenewalClass(renewalStatus)}"
                            >
                                ${safe(renewalStatus)}
                            </strong>
                        </div>

                    </div>

                </section>


                <section class="certificate360-section">

                    <div class="certificate360-section-title">
                        Employee
                    </div>

                    <div class="certificate360-grid">

                        <div class="certificate360-field">
                            <span>Employee Name</span>
                            <strong>
                                ${safe(employeeName)}
                            </strong>
                        </div>

                        <div class="certificate360-field">
                            <span>Employee ID</span>
                            <strong>
                                ${safe(employeeId)}
                            </strong>
                        </div>

                        <div class="certificate360-field">
                            <span>Designation</span>
                            <strong>
                                ${safe(
                                    employee.Designation
                                )}
                            </strong>
                        </div>

                        <div class="certificate360-field">
                            <span>Location</span>
                            <strong>
                                ${safe(
                                    employee.Location
                                )}
                            </strong>
                        </div>

                    </div>

                </section>


                <section class="certificate360-section">

                    <div class="certificate360-section-title">
                        Certificate Document
                    </div>

                    <div class="certificate360-document">

                        <div class="certificate360-document-icon">
                            PDF
                        </div>

                        <div class="certificate360-document-info">

                            <strong>
                                ${safe(documentName)}
                            </strong>

                            <span>
                                Document ID: ${safe(documentId)}
                            </span>

                            <span>
                                Classification: ${safe(classification)}
                            </span>

                        </div>

                        <div class="certificate360-document-action">

                            ${
                                downloadAllowed
                                    ? `
                                        <button
                                            type="button"
                                            class="certificate360-download"
                                            disabled
                                            title="Document download will be enabled through the authorized document service."
                                        >
                                            Download
                                        </button>
                                      `
                                    : `
                                        <span class="certificate360-no-access">
                                            Access Restricted
                                        </span>
                                      `
                            }

                        </div>

                    </div>

                </section>

            </div>


            <footer class="certificate360-footer">

                <span>
                    CCM • Certificate Management
                </span>

                <button
                    type="button"
                    class="certificate360-footer-close"
                    id="certificate360FooterClose"
                >
                    Close
                </button>

            </footer>

        </section>

    `;


    document.body.appendChild(modal);


    /*
       Close modal.
    */

    let closed = false;

    const escapeHandler = event => {

        if (event.key === "Escape") {
            closeModal();
        }

    };


    const closeModal = () => {

        if (closed) {
            return;
        }

        closed = true;

        document.removeEventListener(
            "keydown",
            escapeHandler
        );

        modal.classList.add(
            "certificate360-closing"
        );

        setTimeout(
            () => {

                if (modal.parentNode) {
                    modal.remove();
                }

            },
            150
        );

    };


    const closeButton =
        modal.querySelector(
            "#certificate360Close"
        );


    const footerClose =
        modal.querySelector(
            "#certificate360FooterClose"
        );


    const backdrop =
        modal.querySelector(
            "[data-certificate360-close]"
        );


    if (closeButton) {
        closeButton.addEventListener(
            "click",
            closeModal
        );
    }


    if (footerClose) {
        footerClose.addEventListener(
            "click",
            closeModal
        );
    }


    if (backdrop) {
        backdrop.addEventListener(
            "click",
            closeModal
        );
    }


    document.addEventListener(
        "keydown",
        escapeHandler
    );


    requestAnimationFrame(
        () => {
            modal.classList.add(
                "certificate360-visible"
            );
        }
    );

}

/* =========================================================
   CERTIFICATION EVENTS
========================================================= */

function initializeCertificationOverview() {

    const search =
        document.getElementById(
            "certificateSearch"
        );


    const statusFilter =
        document.getElementById(
            "certificateStatusFilter"
        );


    const departmentFilter =
        document.getElementById(
            "certificateDepartmentFilter"
        );


    const refreshButton =
        document.getElementById(
            "refreshCertificatesButton"
        );


    if (search) {

        search.addEventListener(
            "input",
            renderCertificates
        );

    }


    if (statusFilter) {

        statusFilter.addEventListener(
            "change",
            renderCertificates
        );

    }


    if (departmentFilter) {

        departmentFilter.addEventListener(
            "change",
            renderCertificates
        );

    }


    if (refreshButton) {

        refreshButton.addEventListener(
            "click",
            loadCertificates
        );

    }

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

initializeCertificationOverview();

loadDashboard();
loadCertificates();

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
