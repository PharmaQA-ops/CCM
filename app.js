document.addEventListener("DOMContentLoaded", () => {

    const navItems = document.querySelectorAll(".nav-item");
    const pages = document.querySelectorAll(".page");

    const pageTitle = document.getElementById("pageTitle");

    const sidebar = document.getElementById("sidebar");
    const menuButton = document.getElementById("menuButton");


    /* ==============================
       SIDEBAR NAVIGATION
    ============================== */

    navItems.forEach(item => {

        item.addEventListener("click", () => {

            const page = item.dataset.page;

            navItems.forEach(nav => {
                nav.classList.remove("active");
            });

            item.classList.add("active");


            pages.forEach(section => {
                section.classList.remove("active-page");
            });


            const target = document.getElementById(page + "Page");

            if (target) {
                target.classList.add("active-page");
            }


            const title = item.querySelector("span:last-child");

            if (title) {
                pageTitle.textContent = title.textContent;
            }


            /* Close mobile sidebar */

            sidebar.classList.remove("open");

        });

    });


    /* ==============================
       MOBILE MENU
    ============================== */

    menuButton.addEventListener("click", () => {

        sidebar.classList.toggle("open");

    });


    /* ==============================
       TEST DASHBOARD VALUES
       Temporary UI values only.
       Live API connection comes later.
    ============================== */

    const demoData = {

        total: 5,
        active: 4,
        renewal: 3,
        expired: 1,
        mandatory: 4,
        compliance: 80,

        risk: {
            "0-7": 3,
            "8-15": 0,
            "16-30": 0,
            "31-60": 0,
            "61-90": 0,
            "90+": 1,
            "EXPIRED": 1
        }

    };


    document.getElementById("totalCertifications")
        .textContent = demoData.total;

    document.getElementById("activeCertifications")
        .textContent = demoData.active;

    document.getElementById("renewalCertifications")
        .textContent = demoData.renewal;

    document.getElementById("expiredCertifications")
        .textContent = demoData.expired;

    document.getElementById("mandatoryCertifications")
        .textContent = demoData.mandatory;

    document.getElementById("complianceRate")
        .textContent = demoData.compliance + "%";


    /* ==============================
       EXPIRY RISK BARS
    ============================== */

    const riskValues = Object.values(demoData.risk);

    const maxRisk = Math.max(...riskValues, 1);


    function setRiskBar(barId, valueId, value) {

        const bar = document.getElementById(barId);
        const number = document.getElementById(valueId);

        if (!bar || !number) return;

        number.textContent = value;

        const percentage =
            value === 0
                ? 0
                : Math.max((value / maxRisk) * 100, 4);

        bar.style.width = percentage + "%";

    }


    setRiskBar("risk07", "risk07Value", demoData.risk["0-7"]);
    setRiskBar("risk815", "risk815Value", demoData.risk["8-15"]);
    setRiskBar("risk1630", "risk1630Value", demoData.risk["16-30"]);
    setRiskBar("risk3160", "risk3160Value", demoData.risk["31-60"]);
    setRiskBar("risk6190", "risk6190Value", demoData.risk["61-90"]);
    setRiskBar("risk90", "risk90Value", demoData.risk["90+"]);
    setRiskBar("riskExpired", "riskExpiredValue", demoData.risk["EXPIRED"]);


    /* ==============================
       ADD CERTIFICATE
    ============================== */

    const addButton =
        document.getElementById("addCertificateButton");

    if (addButton) {

        addButton.addEventListener("click", () => {

            alert(
                "Certification creation will be connected to the CCM backend in the next step."
            );

        });

    }

});
