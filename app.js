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

    const certificateDocuments =
    Array.isArray(result?.documents)
        ? result.documents
        : [];

const certificateDocument =
    certificateDocuments[0] || {};

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
                                downloadAllowed && documentId !== "-"
                                    ? `
                                        <button
                                            type="button"
                                            class="certificate360-download"
                                            id="certificate360Download"
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
       11H — SECURE DOCUMENT DOWNLOAD
       Backend remains the authorization authority.
    */

    const downloadButton =
        modal.querySelector(
            "#certificate360Download"
        );


    if (downloadButton) {

        downloadButton.addEventListener(
            "click",
            async () => {

                if (
                    !documentId ||
                    documentId === "-"
                ) {
                    return;
                }


                const originalText =
                    downloadButton.textContent;


                downloadButton.disabled = true;
                downloadButton.textContent =
                    "Authorizing...";


                try {

                    const downloadResult =
                        await ccmExecute(
                            "documentDownload",
                            {
                                documentId:
                                    documentId
                            }
                        );


                    if (
                        !downloadResult ||
                        downloadResult.success !== true ||
                        downloadResult.allowed !== true ||
                        !downloadResult.url
                    ) {
                        throw new Error(
                            downloadResult?.error ||
                            "Document access was denied."
                        );
                    }


                    window.open(
                        downloadResult.url,
                        "_blank",
                        "noopener,noreferrer"
                    );

                } catch (error) {

                    console.error(
                        "CCM document download failed:",
                        error
                    );


                    alert(
                        error.message ||
                        "Unable to open the document."
                    );

                } finally {

                    downloadButton.disabled = false;
                    downloadButton.textContent =
                        originalText;

                }

            }
        );

    }


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


    /*
       11H — Initialize application shell
       after backend authentication succeeds.
    */

    applyNavigationPermissions_();

    initNavigation();
    initMobileMenu();

    initializeCertificationOverview();


    /*
       Load only data permitted for
       the authenticated user.
    */

    if (
        ccmHasPermission_(
            "PERM-DASHBOARD-VIEW"
        )
    ) {
        loadDashboard();
    }


    if (
        ccmHasPermission_(
            "PERM-CERT-VIEW"
        ) ||
        ccmHasPermission_(
            "PERM-CERT-OWN-VIEW"
        )
    ) {
        loadCertificates();
    }


    openAuthorizedDefaultPage_();

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
   11H — RBAC NAVIGATION
========================================================= */

const CCM_PAGE_PERMISSIONS = {

    dashboard: [
        "PERM-DASHBOARD-VIEW"
    ],

    certifications: [
        "PERM-CERT-VIEW",
        "PERM-CERT-OWN-VIEW"
    ],

    employees: [
        "PERM-EMPLOYEE-VIEW"
    ],

    renewals: [
        "PERM-RENEWAL-VIEW"
    ],

    documents: [
        "PERM-DOCUMENT-PUBLIC",
        "PERM-DOCUMENT-INTERNAL",
        "PERM-DOCUMENT-CONFIDENTIAL",
        "PERM-DOCUMENT-RESTRICTED"
    ],

    reports: [
        "PERM-REPORT-VIEW"
    ],

    audit: [
        "PERM-AUDIT-VIEW"
    ],

    settings: [
        "PERM-RBAC-MANAGE"
    ]

};


function ccmHasPermission_(
    permission
) {

    const permissions =
        Array.isArray(
            window.CCM_PERMISSIONS
        )
            ? window.CCM_PERMISSIONS
            : [];


    return permissions.includes(
        permission
    );

}


function ccmHasAnyPermission_(
    permissions
) {

    return (
        Array.isArray(permissions) &&
        permissions.some(
            permission =>
                ccmHasPermission_(
                    permission
                )
        )
    );

}


function isPageAuthorized_(
    page
) {

    const required =
        CCM_PAGE_PERMISSIONS[
            page
        ] || [];


    return (
        required.length === 0 ||
        ccmHasAnyPermission_(
            required
        )
    );

}


function applyNavigationPermissions_() {

    const navItems =
        document.querySelectorAll(
            ".nav-item"
        );


    navItems.forEach(
        item => {

            const page =
                String(
                    item.dataset.page || ""
                ).trim();


            const required =
                CCM_PAGE_PERMISSIONS[
                    page
                ] || [];


            const allowed =
                required.length === 0 ||
                ccmHasAnyPermission_(
                    required
                );


            item.hidden =
                !allowed;


            item.setAttribute(
                "aria-hidden",
                allowed
                    ? "false"
                    : "true"
            );


            if (!allowed) {

                item.classList.remove(
                    "active"
                );

            }

        }
    );

}


/* =========================================================
   CCM — FINAL NAVIGATION
========================================================= */

function navigateToPage_(page, clickedItem = null) {

    page = String(page || "").trim();

    if (!page) {
        console.warn("CCM: Navigation page is empty.");
        return false;
    }

    const target =
        document.getElementById(page + "Page");

    if (!target) {
        console.error(
            "CCM: Page container not found:",
            page + "Page"
        );
        return false;
    }

    /*
       UI navigation.
       Backend remains the real RBAC authority.
    */

    document
        .querySelectorAll(".nav-item")
        .forEach(nav => {
            nav.classList.toggle(
                "active",
                String(nav.dataset.page || "") === page
            );
        });

    document
        .querySelectorAll(".page")
        .forEach(section => {
            section.classList.remove("active-page");
        });

    target.classList.add("active-page");

    /*
       Update breadcrumb.
    */

    const pageTitle =
        document.getElementById("pageTitle");

    const titleSource =
        clickedItem ||
        document.querySelector(
            `.nav-item[data-page="${page}"]`
        );

    const title =
        titleSource?.querySelector(
            "span:last-child"
        );

    if (pageTitle && title) {
        pageTitle.textContent =
            title.textContent.trim();
    }

    /*
       Close mobile sidebar.
    */

    document
        .getElementById("sidebar")
        ?.classList.remove("open");

    /*
       Load page data.
    */

    switch (page) {

        case "dashboard":
            loadDashboard();
            break;

        case "certifications":
            loadCertificationsManagementPage_();
            break;

        case "employees":
            loadEmployeesPage_();
            break;

        case "renewals":
            loadRenewalsPage_();
            break;

        case "documents":
            loadDocumentsPage_();
            break;

        case "reports":
            loadReportsPage_();
            break;

        case "audit":
            loadAuditPage_();
            break;

        case "settings":
            loadSettingsPage_();
            break;

        default:
            console.warn(
                "CCM: Unknown navigation page:",
                page
            );
            return false;
    }

    console.log(
        "CCM: Navigation successful:",
        page
    );

    return true;
}
    /*
       Page-specific live loading.
       Backend RBAC remains authoritative.
    */

    switch (page) {

        case "dashboard":

            if (
                ccmHasPermission_(
                    "PERM-DASHBOARD-VIEW"
                )
            ) {
                loadDashboard();
            }

            break;


        case "certifications":

            if (
                ccmHasPermission_(
                    "PERM-CERT-VIEW"
                ) ||
                ccmHasPermission_(
                    "PERM-CERT-OWN-VIEW"
                )
            ) {
                loadCertificates();
            }

            break;


        default:
            break;

    }


    return true;

}


function openAuthorizedDefaultPage_() {

    const preferredPages = [
        "dashboard",
        "certifications",
        "employees",
        "renewals",
        "documents",
        "reports",
        "audit",
        "settings"
    ];


    const firstAuthorized =
        preferredPages.find(
            page =>
                isPageAuthorized_(
                    page
                )
        );


    if (firstAuthorized) {

        navigateToPage_(
            firstAuthorized
        );

    }

}


function initNavigation() {
    const navItems =
        document.querySelectorAll(".nav-item");

    navItems.forEach(item => {

        // Prevent duplicate handlers
        if (item.dataset.ccmNavigationBound === "true") {
            return;
        }

        item.dataset.ccmNavigationBound = "true";

        item.addEventListener("click", function(event) {

            event.preventDefault();
            event.stopPropagation();

            const page =
                String(
                    item.dataset.page || ""
                ).trim();

            if (!page) {
                console.warn(
                    "CCM navigation item has no data-page:",
                    item
                );
                return;
            }

            console.log(
                "CCM navigation:",
                page
            );

            navigateToPage_(
                page,
                item
            );
        });
    });

    applyNavigationPermissions_();
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


    if (
        menuButton.dataset.ccmMenuBound ===
        "true"
    ) {

        return;

    }


    menuButton.dataset.ccmMenuBound =
        "true";


    menuButton.addEventListener(
        "click",
        () => {

            sidebar.classList.toggle(
                "open"
            );

        }
    );

}

/* =========================================================
   CCM 11I — FULL FUNCTIONAL UI WIRING
   Frontend: GitHub Pages
   Backend: Apps Script scripts.run gateway
========================================================= */

let CCM_EMPLOYEES = [];
let CCM_RENEWALS = [];
let CCM_DOCUMENTS = [];
let CCM_REPORTS = null;
let CCM_AUDIT = [];
let CCM_PAGE_CACHE = {};

function ccmNotify_(message, type = "info") {
    let host = document.getElementById("ccmToastHost");
    if (!host) {
        host = document.createElement("div");
        host.id = "ccmToastHost";
        host.className = "ccm-toast-host";
        document.body.appendChild(host);
    }
    const toast = document.createElement("div");
    toast.className = `ccm-toast ccm-toast-${type}`;
    toast.textContent = message;
    host.appendChild(toast);
    setTimeout(() => toast.classList.add("ccm-toast-show"), 10);
    setTimeout(() => {
        toast.classList.remove("ccm-toast-show");
        setTimeout(() => toast.remove(), 180);
    }, 3200);
}

function ccmFormat_(value) {
    return value === null || value === undefined || value === "" ? "—" : String(value);
}

function ccmStatusBadge_(value) {
    const v = String(value || "—").toUpperCase();
    let cls = "ccm-badge-neutral";
    if (["ACTIVE", "SUCCESS", "RENEWED", "ONLINE"].includes(v)) cls = "ccm-badge-success";
    if (["CRITICAL", "EXPIRED", "FAILED", "DENIED"].includes(v)) cls = "ccm-badge-danger";
    if (["URGENT", "RENEWAL_DUE", "PLANNING", "UPCOMING"].includes(v)) cls = "ccm-badge-warning";
    return `<span class="ccm-badge ${cls}">${escapeHtml(v)}</span>`;
}

function ccmPageShell_(title, eyebrow, description, actionHtml = "") {
    return `
        <div class="page-heading">
            <div>
                <span class="eyebrow">${escapeHtml(eyebrow)}</span>
                <h1>${escapeHtml(title)}</h1>
                <p>${escapeHtml(description)}</p>
            </div>
            ${actionHtml}
        </div>
    `;
}

function ccmPanel_(html, extraClass = "") {
    return `<div class="panel ccm-functional-panel ${extraClass}">${html}</div>`;
}

function ccmLoading_(message = "Loading live data…") {
    return `<div class="ccm-loading-state"><span class="ccm-spinner"></span><span>${escapeHtml(message)}</span></div>`;
}

function ccmError_(message, retryAction = "") {
    return `<div class="ccm-error-state"><strong>Unable to load data</strong><span>${escapeHtml(message)}</span>${retryAction ? `<button type="button" class="primary-button" onclick="${retryAction}">Retry</button>` : ""}</div>`;
}

async function ccmSafeExecute_(action, params = {}) {
    try {
        return await ccmExecute(action, params);
    } catch (error) {
        console.error(`CCM action failed: ${action}`, error);
        throw error;
    }
}

/* =========================================================
   PAGE LOADERS
========================================================= */

async function loadEmployeesPage_() {
    const page = document.getElementById("employeesPage");
    if (!page) return;
    page.innerHTML = ccmPageShell_("Employees", "MASTER DATA", "Employee certification profiles and compliance.") + ccmPanel_(ccmLoading_());
    try {
        const result = await ccmSafeExecute_("employees", {});
        CCM_EMPLOYEES = Array.isArray(result.data) ? result.data : [];
        CCM_PAGE_CACHE.employees = CCM_EMPLOYEES;
        renderEmployeesPage_();
    } catch (error) {
        page.innerHTML = ccmPageShell_("Employees", "MASTER DATA", "Employee certification profiles and compliance.") + ccmPanel_(ccmError_(error.message || "Employees could not be loaded.", "loadEmployeesPage_()"));
    }
}

function renderEmployeesPage_() {
    const page = document.getElementById("employeesPage");
    if (!page) return;
    page.innerHTML = ccmPageShell_("Employees", "MASTER DATA", "Employee certification profiles and compliance.", `<span class="ccm-page-count">${CCM_EMPLOYEES.length} employee${CCM_EMPLOYEES.length === 1 ? "" : "s"}</span>`) + `
        ${ccmPanel_(`
            <div class="ccm-toolbar">
                <input class="ccm-control" id="employeeSearch" type="search" placeholder="Search employee, ID, department or email…" autocomplete="off">
                <button type="button" class="text-button" id="employeeRefreshButton">Refresh ↻</button>
            </div>
            <div class="ccm-table-wrap">
                <table class="ccm-data-table">
                    <thead><tr><th>Employee</th><th>Employee ID</th><th>Department</th><th>Designation</th><th>Location</th><th>Status</th><th>Action</th></tr></thead>
                    <tbody id="employeeTableBody"></tbody>
                </table>
            </div>
        `)}
    `;
    renderEmployeeRows_();
    document.getElementById("employeeSearch")?.addEventListener("input", renderEmployeeRows_);
    document.getElementById("employeeRefreshButton")?.addEventListener("click", loadEmployeesPage_);
}

function renderEmployeeRows_() {
    const body = document.getElementById("employeeTableBody");
    if (!body) return;
    const q = String(document.getElementById("employeeSearch")?.value || "").trim().toLowerCase();
    const rows = CCM_EMPLOYEES.filter(e => [e.Employee_Name, e.Employee_ID, e.Department, e.Designation, e.Location, e.Email].some(v => String(v || "").toLowerCase().includes(q)));
    body.innerHTML = rows.length ? rows.map(e => `
        <tr>
            <td><strong>${escapeHtml(ccmFormat_(e.Employee_Name))}</strong><small class="ccm-muted-block">${escapeHtml(ccmFormat_(e.Email))}</small></td>
            <td class="ccm-mono">${escapeHtml(ccmFormat_(e.Employee_ID))}</td>
            <td>${escapeHtml(ccmFormat_(e.Department))}</td>
            <td>${escapeHtml(ccmFormat_(e.Designation))}</td>
            <td>${escapeHtml(ccmFormat_(e.Location))}</td>
            <td>${ccmStatusBadge_(e.Employment_Status)}</td>
            <td><button type="button" class="table-action" data-employee-id="${escapeHtml(e.Employee_ID)}">View 360</button></td>
        </tr>
    `).join("") : `<tr><td colspan="7" class="table-empty">No employees found.</td></tr>`;
    body.querySelectorAll("[data-employee-id]").forEach(btn => btn.addEventListener("click", () => openEmployee360_(btn.dataset.employeeId)));
}

async function openEmployee360_(employeeId) {
    try {
        const result = await ccmSafeExecute_("employee360", { employeeId });
        showEmployee360Modal_(result);
    } catch (error) {
        ccmNotify_(error.message || "Unable to open employee profile.", "error");
    }
}

function showEmployee360Modal_(result) {
    const employee = result.employee || {};
    const summary = result.summary || {};
    const certs = Array.isArray(result.certifications) ? result.certifications : [];
    ccmOpenModal_("employee360Modal", `
        <div class="ccm-modal-header"><div><span class="eyebrow">EMPLOYEE 360</span><h2>${escapeHtml(ccmFormat_(employee.Employee_Name))}</h2><p>${escapeHtml(ccmFormat_(employee.Email))}</p></div><button type="button" class="ccm-modal-close" data-ccm-close>×</button></div>
        <div class="ccm-modal-body">
            <div class="ccm-summary-grid">
                <div><span>Total Certifications</span><strong>${ccmFormat_(summary.total || 0)}</strong></div>
                <div><span>Active</span><strong>${ccmFormat_(summary.active || 0)}</strong></div>
                <div><span>Expiring ≤90 Days</span><strong>${ccmFormat_(summary.expiring || 0)}</strong></div>
                <div><span>Expired</span><strong>${ccmFormat_(summary.expired || 0)}</strong></div>
            </div>
            ${ccmPanel_(`<div class="ccm-field-grid"><div><span>Employee ID</span><strong>${escapeHtml(ccmFormat_(employee.Employee_ID))}</strong></div><div><span>Department</span><strong>${escapeHtml(ccmFormat_(employee.Department))}</strong></div><div><span>Designation</span><strong>${escapeHtml(ccmFormat_(employee.Designation))}</strong></div><div><span>Location</span><strong>${escapeHtml(ccmFormat_(employee.Location))}</strong></div><div><span>Employment Status</span><strong>${escapeHtml(ccmFormat_(employee.Employment_Status))}</strong></div><div><span>Joining Date</span><strong>${escapeHtml(ccmFormat_(employee.Joining_Date))}</strong></div></div>`, "ccm-inner-panel")}
            ${ccmPanel_(`<h3 class="ccm-section-title">Certification Record</h3><div class="ccm-table-wrap"><table class="ccm-data-table"><thead><tr><th>Certificate</th><th>Certification</th><th>Expiry</th><th>Status</th><th>Renewal</th><th></th></tr></thead><tbody>${certs.length ? certs.map(c => `<tr><td class="ccm-mono">${escapeHtml(ccmFormat_(c.Certificate_ID))}</td><td>${escapeHtml(ccmFormat_(c.Certification_Name))}</td><td>${escapeHtml(ccmFormat_(c.Expiry_Date))}</td><td>${ccmStatusBadge_(c.Status)}</td><td>${ccmStatusBadge_(c.Renewal_Status)}</td><td><button type="button" class="table-action" data-cert-open="${escapeHtml(c.Certificate_ID)}">View</button></td></tr>`).join("") : `<tr><td colspan="6" class="table-empty">No certifications.</td></tr>`}</tbody></table></div>`, "ccm-inner-panel")}
        </div>
        <div class="ccm-modal-footer"><button type="button" class="footer-button" data-ccm-close>Close</button></div>
    `);
    document.querySelectorAll("[data-cert-open]").forEach(b => b.addEventListener("click", () => openCertificate360(b.dataset.certOpen)));
}

async function loadRenewalsPage_() {
    const page = document.getElementById("renewalsPage");
    if (!page) return;
    page.innerHTML = ccmPageShell_("Renewals", "COMPLIANCE CONTROL", "Monitor certification expiry and renewal actions.", `<button type="button" class="primary-button" id="runRenewalCheckButton">Run Renewal Check</button>`) + ccmPanel_(ccmLoading_());
    try {
        const result = await ccmSafeExecute_("renewals", {});
        CCM_RENEWALS = Array.isArray(result.data) ? result.data : [];
        renderRenewalsPage_();
    } catch (error) {
        page.innerHTML = ccmPageShell_("Renewals", "COMPLIANCE CONTROL", "Monitor certification expiry and renewal actions.", `<button type="button" class="primary-button" id="runRenewalCheckButton">Run Renewal Check</button>`) + ccmPanel_(ccmError_(error.message || "Renewals could not be loaded.", "loadRenewalsPage_()"));
        bindRenewalRunButton_();
    }
}

function renderRenewalsPage_() {
    const page = document.getElementById("renewalsPage");
    if (!page) return;
    page.innerHTML = ccmPageShell_("Renewals", "COMPLIANCE CONTROL", "Monitor certification expiry and renewal actions.", `<button type="button" class="primary-button" id="runRenewalCheckButton">Run Renewal Check</button>`) + `
        ${ccmPanel_(`<div class="ccm-toolbar"><input class="ccm-control" id="renewalSearch" type="search" placeholder="Search renewal…"><select class="ccm-control" id="renewalPriorityFilter"><option value="">All Priorities</option><option>CRITICAL</option><option>HIGH</option><option>MEDIUM</option><option>LOW</option></select><button type="button" class="text-button" id="renewalRefreshButton">Refresh ↻</button></div><div class="ccm-table-wrap"><table class="ccm-data-table"><thead><tr><th>Certificate</th><th>Employee</th><th>Certification</th><th>Expiry</th><th>Days</th><th>Priority</th><th>Status</th><th>Action</th></tr></thead><tbody id="renewalTableBody"></tbody></table></div>`)}
    `;
    renderRenewalRows_();
    document.getElementById("renewalSearch")?.addEventListener("input", renderRenewalRows_);
    document.getElementById("renewalPriorityFilter")?.addEventListener("change", renderRenewalRows_);
    document.getElementById("renewalRefreshButton")?.addEventListener("click", loadRenewalsPage_);
    bindRenewalRunButton_();
}

function renderRenewalRows_() {
    const body = document.getElementById("renewalTableBody");
    if (!body) return;
    const q = String(document.getElementById("renewalSearch")?.value || "").toLowerCase().trim();
    const priority = String(document.getElementById("renewalPriorityFilter")?.value || "");
    const rows = CCM_RENEWALS.filter(r => (!priority || String(r.priority) === priority) && [r.certificateId, r.employee, r.department, r.certification, r.expiryDate, r.priority, r.status].some(v => String(v || "").toLowerCase().includes(q)));
    body.innerHTML = rows.length ? rows.map(r => `<tr><td class="ccm-mono">${escapeHtml(ccmFormat_(r.certificateId))}</td><td>${escapeHtml(ccmFormat_(r.employee))}</td><td>${escapeHtml(ccmFormat_(r.certification))}</td><td>${escapeHtml(ccmFormat_(r.expiryDate))}</td><td class="ccm-days ${Number(r.daysRemaining) < 0 ? "ccm-days-danger" : Number(r.daysRemaining) <= 7 ? "ccm-days-warning" : ""}">${escapeHtml(r.daysRemaining)}</td><td>${ccmStatusBadge_(r.priority)}</td><td>${ccmStatusBadge_(r.status)}</td><td><button type="button" class="table-action" data-renewal-cert="${escapeHtml(r.certificateId)}">View</button></td></tr>`).join("") : `<tr><td colspan="8" class="table-empty">No renewal records.</td></tr>`;
    body.querySelectorAll("[data-renewal-cert]").forEach(b => b.addEventListener("click", () => openCertificate360(b.dataset.renewalCert)));
}

function bindRenewalRunButton_() {
    const btn = document.getElementById("runRenewalCheckButton");
    if (!btn || btn.dataset.bound === "true") return;
    btn.dataset.bound = "true";
    if (!ccmHasPermission_("PERM-RENEWAL-MANAGE")) {
        btn.hidden = true;
        return;
    }
    btn.addEventListener("click", runRenewalCheckFromUI_);
}

async function runRenewalCheckFromUI_() {
    const btn = document.getElementById("runRenewalCheckButton");
    if (btn) { btn.disabled = true; btn.textContent = "Processing…"; }
    try {
        const result = await ccmSafeExecute_("runRenewalCheck", {});
        ccmNotify_(`Renewal check complete. Checked ${result.checked || 0}; sent ${result.sent || 0}; skipped ${result.skipped || 0}.`, "success");
        await loadRenewalsPage_();
        if (ccmHasPermission_("PERM-DASHBOARD-VIEW")) await loadDashboard();
    } catch (error) {
        ccmNotify_(error.message || "Renewal check failed.", "error");
        if (btn) { btn.disabled = false; btn.textContent = "Run Renewal Check"; }
    }
}

async function loadDocumentsPage_() {
    const page = document.getElementById("documentsPage");
    if (!page) return;
    page.innerHTML = ccmPageShell_("Documents", "DOCUMENT CONTROL", "Controlled access to certification documents.") + ccmPanel_(ccmLoading_());
    try {
        const result = await ccmSafeExecute_("documents", {});
        CCM_DOCUMENTS = Array.isArray(result.data) ? result.data : [];
        renderDocumentsPage_();
    } catch (error) {
        page.innerHTML = ccmPageShell_("Documents", "DOCUMENT CONTROL", "Controlled access to certification documents.") + ccmPanel_(ccmError_(error.message || "Documents could not be loaded.", "loadDocumentsPage_()"));
    }
}

function renderDocumentsPage_() {
    const page = document.getElementById("documentsPage");
    if (!page) return;
    page.innerHTML = ccmPageShell_("Documents", "DOCUMENT CONTROL", "Controlled access to certification documents.", `<span class="ccm-page-count">${CCM_DOCUMENTS.length} accessible document${CCM_DOCUMENTS.length === 1 ? "" : "s"}</span>`) + `
        ${ccmPanel_(`<div class="ccm-toolbar"><input class="ccm-control" id="documentSearch" type="search" placeholder="Search document or certificate…"><select class="ccm-control" id="documentClassFilter"><option value="">All Classifications</option><option>PUBLIC</option><option>INTERNAL</option><option>CONFIDENTIAL</option><option>RESTRICTED</option><option>EMPLOYEE_PRIVATE</option></select><button type="button" class="text-button" id="documentRefreshButton">Refresh ↻</button></div><div class="ccm-table-wrap"><table class="ccm-data-table"><thead><tr><th>Document</th><th>Certificate</th><th>Classification</th><th>Login</th><th>Status</th><th>Access</th><th>Action</th></tr></thead><tbody id="documentTableBody"></tbody></table></div>`)}
    `;
    renderDocumentRows_();
    document.getElementById("documentSearch")?.addEventListener("input", renderDocumentRows_);
    document.getElementById("documentClassFilter")?.addEventListener("change", renderDocumentRows_);
    document.getElementById("documentRefreshButton")?.addEventListener("click", loadDocumentsPage_);
}

function renderDocumentRows_() {
    const body = document.getElementById("documentTableBody");
    if (!body) return;
    const q = String(document.getElementById("documentSearch")?.value || "").toLowerCase().trim();
    const cls = String(document.getElementById("documentClassFilter")?.value || "");
    const rows = CCM_DOCUMENTS.filter(d => (!cls || String(d.Classification) === cls) && [d.Document_Name, d.Document_ID, d.Certificate_ID, d.Classification, d.Status].some(v => String(v || "").toLowerCase().includes(q)));
    body.innerHTML = rows.length ? rows.map(d => `<tr><td><strong>${escapeHtml(ccmFormat_(d.Document_Name))}</strong><small class="ccm-muted-block ccm-mono">${escapeHtml(ccmFormat_(d.Document_ID))}</small></td><td class="ccm-mono">${escapeHtml(ccmFormat_(d.Certificate_ID))}</td><td>${ccmStatusBadge_(d.Classification)}</td><td>${escapeHtml(String(d.Login_Required || "—"))}</td><td>${ccmStatusBadge_(d.Status)}</td><td>${d.Download_Allowed ? `<span class="ccm-access-ok">Download allowed</span>` : `<span class="ccm-access-denied">Restricted</span>`}</td><td>${d.Download_Allowed ? `<button type="button" class="table-action" data-document-id="${escapeHtml(d.Document_ID)}">Open</button>` : `<span class="ccm-muted-block">No access</span>`}</td></tr>`).join("") : `<tr><td colspan="7" class="table-empty">No accessible documents.</td></tr>`;
    body.querySelectorAll("[data-document-id]").forEach(b => b.addEventListener("click", () => downloadDocument_(b.dataset.documentId)));
}

async function downloadDocument_(documentId) {
    try {
        const result = await ccmSafeExecute_("documentDownload", { documentId });
        if (!result.url) throw new Error("Document URL was not returned.");
        window.open(result.url, "_blank", "noopener,noreferrer");
    } catch (error) {
        ccmNotify_(error.message || "Document access denied.", "error");
    }
}

async function loadReportsPage_() {
    const page = document.getElementById("reportsPage");
    if (!page) return;
    page.innerHTML = ccmPageShell_("Reports", "ANALYTICS", "Certification and compliance reporting.", `<button type="button" class="primary-button" id="reportRefreshButton">Refresh Report ↻</button>`) + ccmPanel_(ccmLoading_());
    try {
        CCM_REPORTS = await ccmSafeExecute_("reports", {});
        renderReportsPage_();
    } catch (error) {
        page.innerHTML = ccmPageShell_("Reports", "ANALYTICS", "Certification and compliance reporting.", `<button type="button" class="primary-button" id="reportRefreshButton">Refresh Report ↻</button>`) + ccmPanel_(ccmError_(error.message || "Reports could not be loaded.", "loadReportsPage_()"));
        document.getElementById("reportRefreshButton")?.addEventListener("click", loadReportsPage_);
    }
}

function ccmChartRows_(obj) {
    const entries = Object.entries(obj || {}).sort((a,b) => b[1] - a[1]);
    const max = Math.max(...entries.map(x => Number(x[1]) || 0), 1);
    return entries.length ? entries.map(([label, value]) => `<div class="ccm-chart-row"><div class="ccm-chart-label"><span>${escapeHtml(label)}</span><strong>${escapeHtml(value)}</strong></div><div class="ccm-chart-track"><div class="ccm-chart-fill" style="width:${Math.round((Number(value) / max) * 100)}%"></div></div></div>`).join("") : `<div class="table-empty">No data.</div>`;
}

function ccmDonut_(obj) {
    const entries = Object.entries(obj || {});
    const total = entries.reduce((n, [,v]) => n + Number(v || 0), 0);
    if (!total) return `<div class="ccm-donut-empty">No data</div>`;
    let start = 0;
    const stops = [];
    const palette = ["#17365D", "#2D6A9F", "#5D8DB8", "#8BAAC7", "#C2D0DC", "#7A8794"];
    entries.forEach(([label, value], i) => {
        const pct = (Number(value || 0) / total) * 100;
        stops.push(`${palette[i % palette.length]} ${start}% ${start + pct}%`);
        start += pct;
    });
    return `<div class="ccm-donut" style="background:conic-gradient(${stops.join(",")})"><div class="ccm-donut-hole"><strong>${total}</strong><span>Total</span></div></div><div class="ccm-legend">${entries.map(([label,value],i) => `<div><i style="background:${palette[i % palette.length]}"></i><span>${escapeHtml(label)}</span><strong>${value}</strong></div>`).join("")}</div>`;
}

function renderReportsPage_() {
    const page = document.getElementById("reportsPage");
    if (!page) return;
    const r = CCM_REPORTS || {};
    page.innerHTML = ccmPageShell_("Reports", "ANALYTICS", "Certification and compliance reporting.", `<button type="button" class="primary-button" id="reportRefreshButton">Refresh Report ↻</button>`) + `
        <div class="ccm-report-kpis"><div class="kpi-card"><div class="kpi-icon">▣</div><div><span>Total Certifications</span><strong>${ccmFormat_(r.total || 0)}</strong></div></div><div class="kpi-card"><div class="kpi-icon">◫</div><div><span>Categories</span><strong>${Object.keys(r.byCategory || {}).length}</strong></div></div><div class="kpi-card"><div class="kpi-icon">⌂</div><div><span>Issuers</span><strong>${Object.keys(r.byIssuer || {}).length}</strong></div></div><div class="kpi-card"><div class="kpi-icon">▤</div><div><span>Departments</span><strong>${Object.keys(r.byDepartment || {}).length}</strong></div></div></div>
        <div class="ccm-report-grid">
            ${ccmPanel_(`<div class="panel-header"><div><h2>By Category</h2><span>Certification distribution</span></div></div><div class="ccm-chart-area">${ccmDonut_(r.byCategory)}</div>`, "ccm-report-panel")}
            ${ccmPanel_(`<div class="panel-header"><div><h2>By Department</h2><span>Department distribution</span></div></div><div class="ccm-chart-area">${ccmChartRows_(r.byDepartment)}</div>`, "ccm-report-panel")}
            ${ccmPanel_(`<div class="panel-header"><div><h2>By Issuing Body</h2><span>Certification issuer distribution</span></div></div><div class="ccm-chart-area">${ccmChartRows_(r.byIssuer)}</div>`, "ccm-report-panel")}
        </div>
    `;
    document.getElementById("reportRefreshButton")?.addEventListener("click", loadReportsPage_);
}

async function loadAuditPage_() {
    const page = document.getElementById("auditPage");
    if (!page) return;
    page.innerHTML = ccmPageShell_("Audit Trail", "GOVERNANCE", "System activity and controlled actions.") + ccmPanel_(ccmLoading_());
    try {
        const result = await ccmSafeExecute_("audit", {});
        CCM_AUDIT = Array.isArray(result.data) ? result.data : [];
        renderAuditPage_();
    } catch (error) {
        page.innerHTML = ccmPageShell_("Audit Trail", "GOVERNANCE", "System activity and controlled actions.") + ccmPanel_(ccmError_(error.message || "Audit records could not be loaded.", "loadAuditPage_()"));
    }
}

function renderAuditPage_() {
    const page = document.getElementById("auditPage");
    if (!page) return;
    page.innerHTML = ccmPageShell_("Audit Trail", "GOVERNANCE", "System activity and controlled actions.", `<button type="button" class="text-button" id="auditRefreshButton">Refresh ↻</button>`) + ccmPanel_(`<div class="ccm-toolbar"><input class="ccm-control" id="auditSearch" type="search" placeholder="Search action, user, module or record…"></div><div class="ccm-table-wrap"><table class="ccm-data-table"><thead><tr><th>Timestamp</th><th>User</th><th>Role</th><th>Action</th><th>Module</th><th>Record</th><th>Description</th><th>Result</th></tr></thead><tbody id="auditTableBody"></tbody></table></div>`);
    renderAuditRows_();
    document.getElementById("auditSearch")?.addEventListener("input", renderAuditRows_);
    document.getElementById("auditRefreshButton")?.addEventListener("click", loadAuditPage_);
}

function renderAuditRows_() {
    const body = document.getElementById("auditTableBody");
    if (!body) return;
    const q = String(document.getElementById("auditSearch")?.value || "").toLowerCase().trim();
    const rows = CCM_AUDIT.filter(a => Object.values(a || {}).some(v => String(v || "").toLowerCase().includes(q))).slice(0, 500);
    body.innerHTML = rows.length ? rows.map(a => `<tr><td>${escapeHtml(ccmFormat_(a.Timestamp || a.Date || a.Created_Date))}</td><td>${escapeHtml(ccmFormat_(a.Email || a.User_Email))}</td><td>${escapeHtml(ccmFormat_(a.Role_ID))}</td><td>${escapeHtml(ccmFormat_(a.Action))}</td><td>${escapeHtml(ccmFormat_(a.Module))}</td><td class="ccm-mono">${escapeHtml(ccmFormat_(a.Record_ID))}</td><td>${escapeHtml(ccmFormat_(a.Description))}</td><td>${ccmStatusBadge_(a.Result)}</td></tr>`).join("") : `<tr><td colspan="8" class="table-empty">No audit records.</td></tr>`;
}

async function loadSettingsPage_() {
    const page = document.getElementById("settingsPage");
    if (!page) return;
    page.innerHTML = ccmPageShell_("Settings", "SYSTEM", "CCM configuration and administration.") + ccmPanel_(ccmLoading_());
    try {
        const result = await ccmSafeExecute_("settings", {});
        renderSettingsPage_(result);
    } catch (error) {
        page.innerHTML = ccmPageShell_("Settings", "SYSTEM", "CCM configuration and administration.") + ccmPanel_(ccmError_(error.message || "Settings could not be loaded.", "loadSettingsPage_()"));
    }
}

function renderSettingsPage_(result) {
    const page = document.getElementById("settingsPage");
    if (!page) return;
    const rows = Array.isArray(result.data) ? result.data : [];
    page.innerHTML = ccmPageShell_("Settings", "SYSTEM", "CCM configuration and administration.") + `
        ${ccmPanel_(`<div class="ccm-admin-banner"><strong>RBAC protected administration</strong><span>Settings are read from the production configuration sheet. Changes should be performed through controlled administration procedures.</span></div><div class="ccm-table-wrap"><table class="ccm-data-table"><thead><tr><th>Setting</th><th>Value</th><th>Status</th></tr></thead><tbody>${rows.length ? rows.map(r => `<tr><td class="ccm-mono">${escapeHtml(ccmFormat_(r.Key || r.Setting || r.Name))}</td><td>${escapeHtml(ccmFormat_(r.Value))}</td><td>${ccmStatusBadge_(r.Status || "CONFIGURED")}</td></tr>`).join("") : `<tr><td colspan="3" class="table-empty">No settings records.</td></tr>`}</tbody></table></div>`)}
        ${ccmPanel_(`<div class="ccm-field-grid"><div><span>Current User</span><strong>${escapeHtml(ccmFormat_(window.CCM_CURRENT_USER?.Email))}</strong></div><div><span>Role</span><strong>${escapeHtml(ccmFormat_(window.CCM_CURRENT_USER?.Role_ID))}</strong></div><div><span>Department</span><strong>${escapeHtml(ccmFormat_(window.CCM_CURRENT_USER?.Department))}</strong></div><div><span>Permissions</span><strong>${Array.isArray(window.CCM_PERMISSIONS) ? window.CCM_PERMISSIONS.length : 0}</strong></div></div>`, "ccm-inner-panel")}
    `;
}

/* =========================================================
   ADD / EDIT CERTIFICATION
========================================================= */

function openAddCertificateModal_() {
    if (!ccmHasPermission_("PERM-CERT-CREATE")) {
        ccmNotify_("You do not have permission to create certifications.", "error");
        return;
    }
    const employees = Array.isArray(CCM_EMPLOYEES) && CCM_EMPLOYEES.length ? CCM_EMPLOYEES : [];
    ccmOpenModal_("certificateFormModal", `
        <div class="ccm-modal-header"><div><span class="eyebrow">CERTIFICATION CONTROL</span><h2>Add Certification</h2><p>Create a controlled certification record in the master register.</p></div><button type="button" class="ccm-modal-close" data-ccm-close>×</button></div>
        <form id="certificateCreateForm" class="ccm-modal-body ccm-form">
            <div class="ccm-form-grid">
                <label>Employee<select name="Employee_ID" id="certificateEmployeeSelect" required><option value="">Select employee</option>${employees.map(e => `<option value="${escapeHtml(e.Employee_ID)}">${escapeHtml(e.Employee_ID + " — " + (e.Employee_Name || ""))}</option>`).join("")}</select></label>
                <label>Employee Name<input name="Employee_Name" id="certificateEmployeeName" readonly></label>
                <label>Department<input name="Department" id="certificateDepartment" readonly></label>
                <label>Designation<input name="Designation" id="certificateDesignation" readonly></label>
                <label>Certification Name<input name="Certification_Name" required placeholder="e.g. ISO 9001 Awareness"></label>
                <label>Certification Category<input name="Certification_Category" placeholder="e.g. Quality / Safety / Security"></label>
                <label>Issuing Body<input name="Issuing_Body" placeholder="Issuing organization"></label>
                <label>Certificate Number<input name="Certificate_Number" placeholder="Certificate number"></label>
                <label>Issue Date<input name="Issue_Date" type="date" required></label>
                <label>Expiry Date<input name="Expiry_Date" type="date" required></label>
                <label>Mandatory<select name="Mandatory"><option value="NO">NO</option><option value="YES">YES</option></select></label>
                <label>Document ID<input name="Document_ID" placeholder="Optional document ID"></label>
                <label class="ccm-form-wide">Remarks<textarea name="Remarks" rows="3" placeholder="Remarks / notes"></textarea></label>
            </div>
            <div id="certificateFormStatus" class="ccm-form-status"></div>
        </form>
        <div class="ccm-modal-footer"><button type="button" class="footer-button" data-ccm-close>Cancel</button><button type="submit" form="certificateCreateForm" class="primary-button" id="saveCertificateButton">Create Certification</button></div>
    `);
    const select = document.getElementById("certificateEmployeeSelect");
    const fill = () => {
        const e = employees.find(x => String(x.Employee_ID) === String(select?.value));
        document.getElementById("certificateEmployeeName").value = e?.Employee_Name || "";
        document.getElementById("certificateDepartment").value = e?.Department || "";
        document.getElementById("certificateDesignation").value = e?.Designation || "";
    };
    select?.addEventListener("change", fill);
    document.getElementById("certificateCreateForm")?.addEventListener("submit", submitCertificateCreate_);
}

async function openAddCertificateModal_() {
    if (!ccmHasPermission_("PERM-CERT-CREATE")) {
        ccmNotify_(
            "You do not have permission to create certifications.",
            "error"
        );
        return;
    }

    let employees =
        Array.isArray(CCM_EMPLOYEES)
            ? CCM_EMPLOYEES
            : [];

    /*
     * Ensure employee master data is available
     * before opening the certification form.
     */
    if (!employees.length) {
        try {
            const result =
                await ccmSafeExecute_(
                    "employees",
                    {}
                );

            employees =
                Array.isArray(result.data)
                    ? result.data
                    : [];

            CCM_EMPLOYEES = employees;
            CCM_PAGE_CACHE.employees =
                employees;

        } catch (error) {
            console.error(
                "CCM employee loading for certificate form failed:",
                error
            );

            ccmNotify_(
                error.message ||
                "Employee master data could not be loaded.",
                "error"
            );

            return;
        }
    }

    ccmOpenModal_(
        "certificateFormModal",
        `
        <div class="ccm-modal-header">
            <div>
                <span class="eyebrow">
                    CERTIFICATION CONTROL
                </span>
                <h2>Add Certification</h2>
                <p>
                    Create a controlled certification record
                    in the master register.
                </p>
            </div>

            <button
                type="button"
                class="ccm-modal-close"
                data-ccm-close
            >×</button>
        </div>

        <form
            id="certificateCreateForm"
            class="ccm-modal-body ccm-form"
        >
            <div class="ccm-form-grid">

                <label>
                    Employee
                    <select
                        name="Employee_ID"
                        id="certificateEmployeeSelect"
                        required
                    >
                        <option value="">
                            Select employee
                        </option>

                        ${employees.map(e => `
                            <option
                                value="${escapeHtml(
                                    e.Employee_ID
                                )}"
                            >
                                ${escapeHtml(
                                    e.Employee_ID +
                                    " — " +
                                    (e.Employee_Name || "")
                                )}
                            </option>
                        `).join("")}
                    </select>
                </label>

                <label>
                    Employee Name
                    <input
                        name="Employee_Name"
                        id="certificateEmployeeName"
                        readonly
                    >
                </label>

                <label>
                    Department
                    <input
                        name="Department"
                        id="certificateDepartment"
                        readonly
                    >
                </label>

                <label>
                    Designation
                    <input
                        name="Designation"
                        id="certificateDesignation"
                        readonly
                    >
                </label>

                <label>
                    Certification Name
                    <input
                        name="Certification_Name"
                        required
                        placeholder="e.g. ISO 9001 Awareness"
                    >
                </label>

                <label>
                    Certification Category
                    <input
                        name="Certification_Category"
                        placeholder="e.g. Quality / Safety / Security"
                    >
                </label>

                <label>
                    Issuing Body
                    <input
                        name="Issuing_Body"
                        placeholder="Issuing organization"
                    >
                </label>

                <label>
                    Certificate Number
                    <input
                        name="Certificate_Number"
                        placeholder="Certificate number"
                    >
                </label>

                <label>
                    Issue Date
                    <input
                        name="Issue_Date"
                        type="date"
                        required
                    >
                </label>

                <label>
                    Expiry Date
                    <input
                        name="Expiry_Date"
                        type="date"
                        required
                    >
                </label>

                <label>
                    Mandatory
                    <select name="Mandatory">
                        <option value="NO">NO</option>
                        <option value="YES">YES</option>
                    </select>
                </label>

                <label>
                    Document ID
                    <input
                        name="Document_ID"
                        placeholder="Optional document ID"
                    >
                </label>

                <label class="ccm-form-wide">
                    Remarks
                    <textarea
                        name="Remarks"
                        rows="3"
                        placeholder="Remarks / notes"
                    ></textarea>
                </label>

            </div>

            <div
                id="certificateFormStatus"
                class="ccm-form-status"
            ></div>
        </form>

        <div class="ccm-modal-footer">
            <button
                type="button"
                class="footer-button"
                data-ccm-close
            >
                Cancel
            </button>

            <button
                type="submit"
                form="certificateCreateForm"
                class="primary-button"
                id="saveCertificateButton"
            >
                Create Certification
            </button>
        </div>
        `
    );

    const select =
        document.getElementById(
            "certificateEmployeeSelect"
        );

    const fill = () => {
        const employee =
            employees.find(
                x =>
                    String(x.Employee_ID) ===
                    String(select?.value)
            );

        document.getElementById(
            "certificateEmployeeName"
        ).value =
            employee?.Employee_Name || "";

        document.getElementById(
            "certificateDepartment"
        ).value =
            employee?.Department || "";

        document.getElementById(
            "certificateDesignation"
        ).value =
            employee?.Designation || "";
    };

    select?.addEventListener(
        "change",
        fill
    );

    document
        .getElementById(
            "certificateCreateForm"
        )
        ?.addEventListener(
            "submit",
            submitCertificateCreate_
        );
}
async function submitCertificateCreate_(event) {
    event.preventDefault();

    const form = event.currentTarget;
    const button = document.getElementById("saveCertificateButton");
    const status = document.getElementById("certificateFormStatus");

    if (!form) {
        return;
    }

    const data = Object.fromEntries(
        new FormData(form).entries()
    );

    /*
     * Basic client-side validation
     */
    if (!data.Employee_ID) {
        if (status) {
            status.textContent = "Please select an employee.";
        }
        return;
    }

    if (!data.Certification_Name) {
        if (status) {
            status.textContent = "Certification Name is required.";
        }
        return;
    }

    if (!data.Issue_Date) {
        if (status) {
            status.textContent = "Issue Date is required.";
        }
        return;
    }

    if (!data.Expiry_Date) {
        if (status) {
            status.textContent = "Expiry Date is required.";
        }
        return;
    }

    /*
     * Prevent duplicate submission
     */
    if (button) {
        button.disabled = true;
        button.textContent = "Creating…";
    }

    if (status) {
        status.textContent = "Creating certification record…";
    }

    try {
        const result = await ccmSafeExecute_(
            "createCertificate",
            {
                data: data
            }
        );

        console.log(
            "CCM certification creation result:",
            result
        );

        if (!result || result.success !== true) {
            throw new Error(
                result?.error ||
                "Certification could not be created."
            );
        }

        /*
         * Close form only after backend confirms success.
         */
        ccmCloseModal_(
            "certificateFormModal"
        );

        ccmNotify_(
            `Certification ${result.certificateId || ""} created successfully.`,
            "success"
        );

        /*
         * Refresh live certification data.
         */
        await loadCertificates();

        /*
         * Refresh dashboard KPIs.
         */
        if (
            ccmHasPermission_(
                "PERM-DASHBOARD-VIEW"
            )
        ) {
            await loadDashboard();
        }

        /*
         * Refresh certification management page
         * if currently visible.
         */
        if (
            document
                .getElementById("certificationsPage")
                ?.classList
                .contains("active-page")
        ) {
            await loadCertificationsManagementPage_();
        }

    } catch (error) {

        console.error(
            "CCM certification creation failed:",
            error
        );

        if (status) {
            status.textContent =
                error?.message ||
                "Certification creation failed.";
        }

        if (button) {
            button.disabled = false;
            button.textContent =
                "Create Certification";
        }
    }
}
async function openEditCertificateModal_(certificateId) {
    if (!ccmHasPermission_("PERM-CERT-EDIT")) {
        ccmNotify_("You do not have permission to edit certifications.", "error");
        return;
    }
    const c = CCM_CERTIFICATES.find(x => String(x.Certificate_ID) === String(certificateId));
    if (!c) { ccmNotify_("Certificate record is not loaded.", "error"); return; }
    ccmOpenModal_("certificateEditModal", `
        <div class="ccm-modal-header"><div><span class="eyebrow">CERTIFICATION CONTROL</span><h2>Edit Certification</h2><p>${escapeHtml(c.Certificate_ID)}</p></div><button type="button" class="ccm-modal-close" data-ccm-close>×</button></div>
        <form id="certificateEditForm" class="ccm-modal-body ccm-form"><input type="hidden" name="Certificate_ID" value="${escapeHtml(c.Certificate_ID)}"><div class="ccm-form-grid">
            <label>Certification Name<input name="Certification_Name" required value="${escapeHtml(c.Certification_Name || "")}"></label>
            <label>Category<input name="Certification_Category" value="${escapeHtml(c.Certification_Category || "")}"></label>
            <label>Issuing Body<input name="Issuing_Body" value="${escapeHtml(c.Issuing_Body || "")}"></label>
            <label>Certificate Number<input name="Certificate_Number" value="${escapeHtml(c.Certificate_Number || "")}"></label>
            <label>Issue Date<input name="Issue_Date" type="date" value="${escapeHtml(ccmIsoDate_(c.Issue_Date))}"></label>
            <label>Expiry Date<input name="Expiry_Date" type="date" required value="${escapeHtml(ccmIsoDate_(c.Expiry_Date))}"></label>
            <label>Mandatory<select name="Mandatory"><option value="NO" ${String(c.Mandatory).toUpperCase() === "NO" ? "selected" : ""}>NO</option><option value="YES" ${String(c.Mandatory).toUpperCase() === "YES" ? "selected" : ""}>YES</option></select></label>
            <label>Status<select name="Status"><option value="ACTIVE" ${String(c.Status).toUpperCase() === "ACTIVE" ? "selected" : ""}>ACTIVE</option><option value="INACTIVE" ${String(c.Status).toUpperCase() === "INACTIVE" ? "selected" : ""}>INACTIVE</option></select></label>
            <label class="ccm-form-wide">Remarks<textarea name="Remarks" rows="3">${escapeHtml(c.Remarks || "")}</textarea></label>
        </div><div id="certificateEditStatus" class="ccm-form-status"></div></form>
        <div class="ccm-modal-footer"><button type="button" class="footer-button" data-ccm-close>Cancel</button><button type="submit" form="certificateEditForm" class="primary-button" id="updateCertificateButton">Save Changes</button></div>
    `);
    document.getElementById("certificateEditForm")?.addEventListener("submit", submitCertificateEdit_);
}

function ccmIsoDate_(value) {
    if (!value) return "";
    const m = String(value).match(/^(\d{2})-([A-Za-z]{3})-(\d{4})$/);
    if (m) {
        const months = {Jan:"01",Feb:"02",Mar:"03",Apr:"04",May:"05",Jun:"06",Jul:"07",Aug:"08",Sep:"09",Oct:"10",Nov:"11",Dec:"12"};
        return `${m[3]}-${months[m[2]] || "01"}-${m[1]}`;
    }
    const d = new Date(value);
    return isNaN(d.getTime()) ? "" : d.toISOString().slice(0,10);
}

async function submitCertificateEdit_(event) {
    event.preventDefault();
    const form = event.currentTarget;
    const btn = document.getElementById("updateCertificateButton");
    const status = document.getElementById("certificateEditStatus");
    const data = Object.fromEntries(new FormData(form).entries());
    const id = data.Certificate_ID;
    delete data.Certificate_ID;
    if (btn) { btn.disabled = true; btn.textContent = "Saving…"; }
    try {
        const result = await ccmSafeExecute_("updateCertificate", { certificateId: id, data });
        ccmCloseModal_("certificateEditModal");
        ccmNotify_(`Certification ${result.certificateId || id} updated successfully.`, "success");
        await loadCertificates();
        if (ccmHasPermission_("PERM-DASHBOARD-VIEW")) await loadDashboard();
        if (document.getElementById("certificationsPage")?.classList.contains("active-page")) await loadCertificationsManagementPage_();
    } catch (error) {
        if (status) status.textContent = error.message || "Update failed.";
        if (btn) { btn.disabled = false; btn.textContent = "Save Changes"; }
    }
}

/* =========================================================
   CERTIFICATION MANAGEMENT PAGE
========================================================= */

async function loadCertificationsManagementPage_() {
    const page = document.getElementById("certificationsPage");
    if (!page) return;
    page.innerHTML = ccmPageShell_("Certifications", "MASTER DATA", "Manage and review certification records.", `${ccmHasPermission_("PERM-CERT-CREATE") ? `<button type="button" class="primary-button" id="certPageAddButton">+ Add Certification</button>` : ""}`) + ccmPanel_(ccmLoading_());
    try {
        await loadCertificates();
        renderCertificationsManagementPage_();
    } catch (error) {
        page.innerHTML = ccmPageShell_("Certifications", "MASTER DATA", "Manage and review certification records.") + ccmPanel_(ccmError_(error.message || "Certificates could not be loaded.", "loadCertificationsManagementPage_()"));
    }
}

function renderCertificationsManagementPage_() {
    const page = document.getElementById("certificationsPage");
    if (!page) return;
    page.innerHTML = ccmPageShell_("Certifications", "MASTER DATA", "Manage and review certification records.", `${ccmHasPermission_("PERM-CERT-CREATE") ? `<button type="button" class="primary-button" id="certPageAddButton">+ Add Certification</button>` : ""}`) + `
        ${ccmPanel_(`<div class="ccm-toolbar"><input class="ccm-control" id="certMgmtSearch" type="search" placeholder="Search certification…"><select class="ccm-control" id="certMgmtStatus"><option value="">All Status</option><option>ACTIVE</option><option>EXPIRED</option><option>INACTIVE</option></select><button type="button" class="text-button" id="certMgmtRefresh">Refresh ↻</button></div><div class="ccm-table-wrap"><table class="ccm-data-table"><thead><tr><th>Certificate</th><th>Employee</th><th>Certification</th><th>Department</th><th>Issue</th><th>Expiry</th><th>Status</th><th>Renewal</th><th>Action</th></tr></thead><tbody id="certMgmtBody"></tbody></table></div>`)}
    `;
    document.getElementById("certPageAddButton")?.addEventListener("click", openAddCertificateModal_);
    document.getElementById("certMgmtSearch")?.addEventListener("input", renderCertManagementRows_);
    document.getElementById("certMgmtStatus")?.addEventListener("change", renderCertManagementRows_);
    document.getElementById("certMgmtRefresh")?.addEventListener("click", async () => { await loadCertificates(); renderCertificationsManagementPage_(); });
    renderCertManagementRows_();
}

function renderCertManagementRows_() {
    const body = document.getElementById("certMgmtBody");
    if (!body) return;
    const q = String(document.getElementById("certMgmtSearch")?.value || "").toLowerCase().trim();
    const status = String(document.getElementById("certMgmtStatus")?.value || "");
    const rows = CCM_CERTIFICATES.filter(c => (!status || String(c.Status).toUpperCase() === status) && [c.Certificate_ID,c.Employee_Name,c.Certification_Name,c.Department,c.Certificate_Number].some(v => String(v || "").toLowerCase().includes(q)));
    body.innerHTML = rows.length ? rows.map(c => `<tr><td class="ccm-mono">${escapeHtml(ccmFormat_(c.Certificate_ID))}</td><td>${escapeHtml(ccmFormat_(c.Employee_Name))}</td><td>${escapeHtml(ccmFormat_(c.Certification_Name))}</td><td>${escapeHtml(ccmFormat_(c.Department))}</td><td>${escapeHtml(ccmFormat_(c.Issue_Date))}</td><td>${escapeHtml(ccmFormat_(c.Expiry_Date))}</td><td>${ccmStatusBadge_(c.Status)}</td><td>${ccmStatusBadge_(c.Renewal_Status)}</td><td><button type="button" class="table-action" data-mgmt-view="${escapeHtml(c.Certificate_ID)}">View</button>${ccmHasPermission_("PERM-CERT-EDIT") ? ` <button type="button" class="table-action" data-mgmt-edit="${escapeHtml(c.Certificate_ID)}">Edit</button>` : ""}</td></tr>`).join("") : `<tr><td colspan="9" class="table-empty">No certification records.</td></tr>`;
    body.querySelectorAll("[data-mgmt-view]").forEach(b => b.addEventListener("click", () => openCertificate360(b.dataset.mgmtView)));
    body.querySelectorAll("[data-mgmt-edit]").forEach(b => b.addEventListener("click", () => openEditCertificateModal_(b.dataset.mgmtEdit)));
}

/* =========================================================
   DASHBOARD ACTIONS
========================================================= */

function wireDashboardActions_() {
    const add = document.getElementById("addCertificateButton");
    if (add && add.dataset.bound !== "true") {
        add.dataset.bound = "true";
        add.addEventListener("click", openAddCertificateModal_);
    }
    const overviewRefresh = document.getElementById("refreshCertificatesButton");
    if (overviewRefresh && overviewRefresh.dataset.bound !== "true") {
        overviewRefresh.dataset.bound = "true";
        overviewRefresh.addEventListener("click", loadCertificates);
    }
    const viewAll = document.getElementById("viewAllCertificatesButton");
    if (viewAll && viewAll.dataset.bound !== "true") {
        viewAll.dataset.bound = "true";
        viewAll.addEventListener("click", () => navigateToPage_("certifications"));
    }
    const dashboard = document.getElementById("dashboardPage");
    if (!dashboard || dashboard.dataset.actionsBound === "true") return;
    dashboard.dataset.actionsBound = "true";
    const panelButtons = dashboard.querySelectorAll(".quick-action");
    if (panelButtons[0]) panelButtons[0].addEventListener("click", openAddCertificateModal_);
    if (panelButtons[1]) panelButtons[1].addEventListener("click", () => navigateToPage_("renewals"));
    if (panelButtons[2]) panelButtons[2].addEventListener("click", () => navigateToPage_("documents"));
    if (panelButtons[3]) panelButtons[3].addEventListener("click", () => navigateToPage_("reports"));
    const riskView = dashboard.querySelector(".dashboard-grid .panel:first-child .text-button");
    if (riskView && riskView.dataset.bound !== "true") {
        riskView.dataset.bound = "true";
        riskView.addEventListener("click", () => navigateToPage_("renewals"));
    }
}

/* =========================================================
   MODAL ENGINE
========================================================= */

function ccmOpenModal_(id, html) {
    ccmCloseAllModals_();
    const modal = document.createElement("div");
    modal.id = id;
    modal.className = "ccm-functional-modal";
    modal.innerHTML = `<div class="ccm-functional-backdrop" data-ccm-close></div><section class="ccm-functional-dialog" role="dialog" aria-modal="true">${html}</section>`;
    document.body.appendChild(modal);
    modal.querySelectorAll("[data-ccm-close]").forEach(el => el.addEventListener("click", () => ccmCloseModal_(id)));
    requestAnimationFrame(() => modal.classList.add("ccm-functional-visible"));
}

function ccmCloseModal_(id) {
    const modal = document.getElementById(id);
    if (!modal) return;
    modal.classList.remove("ccm-functional-visible");
    setTimeout(() => modal.remove(), 160);
}

function ccmCloseAllModals_() {
    document.querySelectorAll(".ccm-functional-modal").forEach(m => m.remove());
}

document.addEventListener("keydown", event => {
    if (event.key === "Escape") {
        document.querySelectorAll(".ccm-functional-modal").forEach(m => m.remove());
    }
});

/* =========================================================
   CERTIFICATE DOWNLOAD — bridge current 360 modal
========================================================= */

function wireCertificate360Download_(result) {
    const doc = Array.isArray(result?.documents) ? result.documents[0] : null;
    const button = document.querySelector(".certificate360-download");
    if (!button || !doc || !doc.Document_ID || !doc.Download_Allowed) return;
    button.disabled = false;
    button.style.cursor = "pointer";
    button.style.opacity = "1";
    button.dataset.documentId = doc.Document_ID;
    button.title = "Open authorized document";
    button.addEventListener("click", () => downloadDocument_(doc.Document_ID));
}

/* =========================================================
   NAVIGATION OVERRIDE
========================================================= */

function navigateToPage_(page, clickedItem = null) {
    if (!page || !isPageAuthorized_(page)) {
        ccmNotify_("You are not authorized to access this section.", "error");
        return false;
    }
    const target = document.getElementById(page + "Page");
    if (!target) return false;
    document.querySelectorAll(".nav-item").forEach(nav => nav.classList.toggle("active", nav.dataset.page === page));
    document.querySelectorAll(".page").forEach(section => section.classList.remove("active-page"));
    target.classList.add("active-page");
    const pageTitle = document.getElementById("pageTitle");
    const titleSource = clickedItem || document.querySelector(`.nav-item[data-page="${page}"]`);
    const title = titleSource?.querySelector("span:last-child");
    if (pageTitle && title) pageTitle.textContent = title.textContent.trim();
    document.getElementById("sidebar")?.classList.remove("open");
    switch (page) {
        case "dashboard": if (ccmHasPermission_("PERM-DASHBOARD-VIEW")) loadDashboard(); break;
        case "certifications": loadCertificationsManagementPage_(); break;
        case "employees": loadEmployeesPage_(); break;
        case "renewals": loadRenewalsPage_(); break;
        case "documents": loadDocumentsPage_(); break;
        case "reports": loadReportsPage_(); break;
        case "audit": loadAuditPage_(); break;
        case "settings": loadSettingsPage_(); break;
    }
    return true;
}

function openAuthorizedDefaultPage_() {
    const pages = ["dashboard", "certifications", "employees", "renewals", "documents", "reports", "audit", "settings"];
    const page = pages.find(isPageAuthorized_);
    if (page) navigateToPage_(page);
}

function wireAllCCMUI_() {
    applyNavigationPermissions_();
    initNavigation();
    initMobileMenu();
    initializeCertificationOverview();
    wireDashboardActions_();
    document.querySelectorAll(".icon-button[title=\"Notifications\"]").forEach(btn => {
        if (btn.dataset.bound === "true") return;
        btn.dataset.bound = "true";
        btn.addEventListener("click", () => {
            if (ccmHasPermission_("PERM-RENEWAL-VIEW")) navigateToPage_("renewals");
            else ccmNotify_("No renewal notifications are currently available.", "info");
        });
    });
}

function openCCMApplication() {
    document.getElementById("loginScreen")?.style.setProperty("display", "none");
    wireAllCCMUI_();
    if (ccmHasPermission_("PERM-DASHBOARD-VIEW")) loadDashboard();
    if (ccmHasPermission_("PERM-CERT-VIEW") || ccmHasPermission_("PERM-CERT-OWN-VIEW")) loadCertificates();
    wireDashboardActions_();
    openAuthorizedDefaultPage_();
}

/* =========================================================
   EXTEND CURRENT CERTIFICATE 360 OPEN FUNCTION
========================================================= */

const CCM_ORIGINAL_OPEN_CERTIFICATE_360 = openCertificate360;
openCertificate360 = async function(certificateId) {
    try {
        const result = await ccmSafeExecute_("certificate360", { certificateId });
        showCertificateDetails(result);
        setTimeout(() => wireCertificate360Download_(result), 20);
    } catch (error) {
        ccmNotify_(error.message || "Unable to open certificate.", "error");
    }
};

/* =========================================================
   STARTUP PATCH
========================================================= */

document.addEventListener("DOMContentLoaded", () => {
    setTimeout(() => {
        if (window.CCM_CURRENT_USER) wireAllCCMUI_();
    }, 100);
});

/* =========================================================
   CCM 11I — FINAL ACTION PERMISSION POLISH
========================================================= */

function wireDashboardActionsFinal_() {
    const dashboard = document.getElementById("dashboardPage");
    if (!dashboard) return;

    const addAllowed = ccmHasPermission_("PERM-CERT-CREATE");
    const renewalAllowed = ccmHasPermission_("PERM-RENEWAL-VIEW");
    const documentAllowed = ccmHasAnyPermission_([
        "PERM-DOCUMENT-PUBLIC",
        "PERM-DOCUMENT-INTERNAL",
        "PERM-DOCUMENT-CONFIDENTIAL",
        "PERM-DOCUMENT-RESTRICTED"
    ]);
    const reportAllowed = ccmHasPermission_("PERM-REPORT-VIEW");

    const topAdd = document.getElementById("addCertificateButton");
    if (topAdd) topAdd.hidden = !addAllowed;

    const actions = dashboard.querySelectorAll(".quick-action");
    if (actions[0]) actions[0].hidden = !addAllowed;
    if (actions[1]) actions[1].hidden = !renewalAllowed;
    if (actions[2]) actions[2].hidden = !documentAllowed;
    if (actions[3]) actions[3].hidden = !reportAllowed;
}

function wireProfileMenu_() {
    const button = document.querySelector(".profile-arrow");
    if (!button || button.dataset.bound === "true") return;
    button.dataset.bound = "true";
    button.addEventListener("click", () => {
        const user = window.CCM_CURRENT_USER || {};
        const permissions = Array.isArray(window.CCM_PERMISSIONS) ? window.CCM_PERMISSIONS : [];
        ccmOpenModal_("profileModal", `
            <div class="ccm-modal-header"><div><span class="eyebrow">CCM ACCOUNT</span><h2>${escapeHtml(ccmFormat_(user.Name || user.Email))}</h2><p>${escapeHtml(ccmFormat_(user.Email))}</p></div><button type="button" class="ccm-modal-close" data-ccm-close>×</button></div>
            <div class="ccm-modal-body">
                ${ccmPanel_(`<div class="ccm-field-grid"><div><span>User ID</span><strong>${escapeHtml(ccmFormat_(user.User_ID))}</strong></div><div><span>Employee ID</span><strong>${escapeHtml(ccmFormat_(user.Employee_ID))}</strong></div><div><span>Department</span><strong>${escapeHtml(ccmFormat_(user.Department))}</strong></div><div><span>Role</span><strong>${escapeHtml(ccmFormat_(user.Role_ID))}</strong></div><div><span>Status</span><strong>${escapeHtml(ccmFormat_(user.Status))}</strong></div><div><span>Permissions</span><strong>${permissions.length}</strong></div></div>`, "ccm-inner-panel")}
            </div>
            <div class="ccm-modal-footer"><button type="button" class="footer-button" data-ccm-close>Close</button></div>
        `);
    });
}

function wireAllCCMUI_() {
    applyNavigationPermissions_();
    initNavigation();
    initMobileMenu();
    initializeCertificationOverview();
    wireDashboardActions_();
    wireDashboardActionsFinal_();
    wireProfileMenu_();

    document.querySelectorAll(".icon-button[title=\"Notifications\"]").forEach(btn => {
        if (btn.dataset.bound === "true") return;
        btn.dataset.bound = "true";
        btn.addEventListener("click", () => {
            if (ccmHasPermission_("PERM-RENEWAL-VIEW")) {
                navigateToPage_("renewals");
            } else {
                ccmNotify_("No renewal notifications are available for your account.", "info");
            }
        });
    });
}

function openCCMApplication() {
    const loginScreen = document.getElementById("loginScreen");
    if (loginScreen) loginScreen.style.display = "none";
    wireAllCCMUI_();
    if (ccmHasPermission_("PERM-DASHBOARD-VIEW")) loadDashboard();
    if (ccmHasPermission_("PERM-CERT-VIEW") || ccmHasPermission_("PERM-CERT-OWN-VIEW")) loadCertificates();
    openAuthorizedDefaultPage_();
    setTimeout(wireDashboardActionsFinal_, 250);
}
