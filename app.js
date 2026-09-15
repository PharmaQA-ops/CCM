const CCM_API =
    "https://script.google.com/macros/s/AKfycbyzpJ6vSFQvZXLVxOEpLFuIpB8oQnegtHpKJaSWZll0gQkX6K5FjFAt4W3ugbRYjQOafw/exec";


document.addEventListener("DOMContentLoaded", () => {

    initNavigation();
    initMobileMenu();

    loadCurrentUser();
    loadDashboard();

});


/* =========================================================
   API
========================================================= */

async function ccmApi(action, params = {}) {

    const url = new URL(CCM_API);

    url.searchParams.set("action", action);

    Object.keys(params).forEach(key => {

        if (
            params[key] !== undefined &&
            params[key] !== null &&
            params[key] !== ""
        ) {
            url.searchParams.set(key, params[key]);
        }

    });

    const response = await fetch(url.toString(), {
        method: "GET",
        credentials: "include"
    });

    if (!response.ok) {
        throw new Error(
            `HTTP ${response.status}`
        );
    }

    const data = await response.json();

    if (!data.success) {
        throw new Error(
            data.error || "CCM API request failed"
        );
    }

    return data;
}


/* =========================================================
   CURRENT USER
========================================================= */

async function loadCurrentUser() {

    const userName =
        document.getElementById("userName");

    const userRole =
        document.getElementById("userRole");

    try {

        const data =
            await ccmApi("me");

        if (
            data.authenticated &&
            data.authorized &&
            data.user
        ) {

            userName.textContent =
                data.user.Name || "User";

            userRole.textContent =
                formatRole(data.user.Role_ID);

            updateAvatar(
                data.user.Name || "U"
            );

        } else {

            userName.textContent = "Unauthorized";
            userRole.textContent = "Access Denied";

        }

    } catch (error) {

        console.error(
            "CCM user error:",
            error
        );

        userName.textContent = "User";
        userRole.textContent = "Authentication";

    }

}


/* =========================================================
   DASHBOARD
========================================================= */

async function loadDashboard() {

    setDashboardLoading();

    try {

        const data =
            await ccmApi("dashboard");

        if (!data.kpi) {
            throw new Error(
                "Invalid dashboard response"
            );
        }

        updateKpis(data.kpi);

        updateExpiryRisk(
            data.expiryRisk || {}
        );

    } catch (error) {

        console.error(
            "CCM dashboard error:",
            error
        );

        showDashboardError(
            error.message
        );

    }

}


/* =========================================================
   KPI
========================================================= */

function updateKpis(kpi) {

    setText(
        "totalCertifications",
        valueOrZero(kpi.total)
    );

    setText(
        "activeCertifications",
        valueOrZero(kpi.active)
    );

    setText(
        "renewalCertifications",
        valueOrZero(kpi.renewal)
    );

    setText(
        "expiredCertifications",
        valueOrZero(kpi.expired)
    );

    setText(
        "mandatoryCertifications",
        valueOrZero(kpi.mandatory)
    );

    setText(
        "complianceRate",
        valueOrZero(kpi.compliance) + "%"
    );

}


/* =========================================================
   EXPIRY RISK
========================================================= */

function updateExpiryRisk(risk) {

    const values = [

        Number(risk["0-7"]) || 0,
        Number(risk["8-15"]) || 0,
        Number(risk["16-30"]) || 0,
        Number(risk["31-60"]) || 0,
        Number(risk["61-90"]) || 0,
        Number(risk["90+"]) || 0,
        Number(risk["EXPIRED"]) || 0

    ];

    const maxValue =
        Math.max(...values, 1);


    setRiskBar(
        "risk07",
        "risk07Value",
        values[0],
        maxValue
    );

    setRiskBar(
        "risk815",
        "risk815Value",
        values[1],
        maxValue
    );

    setRiskBar(
        "risk1630",
        "risk1630Value",
        values[2],
        maxValue
    );

    setRiskBar(
        "risk3160",
        "risk3160Value",
        values[3],
        maxValue
    );

    setRiskBar(
        "risk6190",
        "risk6190Value",
        values[4],
        maxValue
    );

    setRiskBar(
        "risk90",
        "risk90Value",
        values[5],
        maxValue
    );

    setRiskBar(
        "riskExpired",
        "riskExpiredValue",
        values[6],
        maxValue
    );

}


function setRiskBar(
    barId,
    valueId,
    value,
    maxValue
) {

    const bar =
        document.getElementById(barId);

    const valueElement =
        document.getElementById(valueId);

    if (!bar || !valueElement) {
        return;
    }

    valueElement.textContent =
        value;

    if (value === 0) {

        bar.style.width = "0%";

        return;
    }

    const percentage =
        Math.max(
            (value / maxValue) * 100,
            4
        );

    bar.style.width =
        percentage + "%";

}


/* =========================================================
   NAVIGATION
========================================================= */

function initNavigation() {

    const navItems =
        document.querySelectorAll(".nav-item");

    const pages =
        document.querySelectorAll(".page");

    const pageTitle =
        document.getElementById("pageTitle");


    navItems.forEach(item => {

        item.addEventListener("click", () => {

            const page =
                item.dataset.page;

            navItems.forEach(nav => {

                nav.classList.remove(
                    "active"
                );

            });

            item.classList.add("active");


            pages.forEach(section => {

                section.classList.remove(
                    "active-page"
                );

            });


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

            if (title) {

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

        });

    });

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

    if (!menuButton || !sidebar) {
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


/* =========================================================
   UI HELPERS
========================================================= */

function setText(id, value) {

    const element =
        document.getElementById(id);

    if (element) {
        element.textContent = value;
    }

}


function valueOrZero(value) {

    if (
        value === undefined ||
        value === null ||
        value === ""
    ) {
        return 0;
    }

    return value;

}


function updateAvatar(name) {

    const avatar =
        document.querySelector(".avatar");

    if (!avatar) {
        return;
    }

    const cleanName =
        String(name).trim();

    if (!cleanName) {
        avatar.textContent = "U";
        return;
    }

    const parts =
        cleanName.split(/\s+/);

    if (parts.length === 1) {

        avatar.textContent =
            parts[0].charAt(0).toUpperCase();

    } else {

        avatar.textContent =
            (
                parts[0].charAt(0) +
                parts[parts.length - 1].charAt(0)
            ).toUpperCase();

    }

}


function formatRole(role) {

    if (!role) {
        return "User";
    }

    return String(role)
        .replace(/^ROLE-/i, "")
        .replace(/[-_]/g, " ")
        .replace(/\b\w/g, char =>
            char.toUpperCase()
        );

}


/* =========================================================
   LOADING
========================================================= */

function setDashboardLoading() {

    const ids = [

        "totalCertifications",
        "activeCertifications",
        "renewalCertifications",
        "expiredCertifications",
        "mandatoryCertifications",
        "complianceRate"

    ];

    ids.forEach(id => {

        setText(id, "…");

    });

}


/* =========================================================
   API ERROR
========================================================= */

function showDashboardError(message) {

    const ids = [

        "totalCertifications",
        "activeCertifications",
        "renewalCertifications",
        "expiredCertifications",
        "mandatoryCertifications"

    ];

    ids.forEach(id => {

        setText(id, "—");

    });

    setText(
        "complianceRate",
        "—"
    );

    console.error(
        "CCM Dashboard API Error:",
        message
    );

}
