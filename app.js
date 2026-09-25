/************************************************************
 * GGL CENTRAL CERTIFICATION MANAGER
 * PRODUCTION BACKEND
 *
 * Backend:
 * Google Apps Script
 * Google Sheets
 * Google Drive
 * Gmail
 *
 * Frontend:
 * GitHub Pages
 *
 * VERSION: 12.0.0
 ************************************************************/

const CONFIG = {
  COMPANY: "Gaerish Logistics Pvt Ltd",
  SYSTEM: "GGL Central Certification Manager",
  TIMEZONE: "Asia/Kolkata",

  SHEETS: {
    CERTIFICATIONS: "01_CERTIFICATIONS",
    EMPLOYEES: "02_EMPLOYEES",
    TYPES: "03_CERTIFICATION_TYPES",
    DEPARTMENTS: "04_DEPARTMENTS",
    RENEWAL_RULES: "05_RENEWAL_RULES",
    USERS: "06_USERS",
    ROLES: "07_ROLES",
    PERMISSIONS: "08_PERMISSIONS",
    ROLE_PERMISSIONS: "09_ROLE_PERMISSIONS",
    DOCUMENTS: "10_DOCUMENTS",
    NOTIFICATION_LOG: "11_NOTIFICATION_LOG",
    AUDIT_LOG: "12_AUDIT_LOG",
    SETTINGS: "13_SETTINGS",
    IMPORT_EXPORT_AUTHORIZATIONS: "14_IMPORT_EXPORT_AUTHORIZATIONS"
  }
};

var CCM_REQUEST_USER_EMAIL = "";


/************************************************************
 * WEB APP ENTRY
 ************************************************************/

function doGet(e) {

  const action =
    e &&
    e.parameter &&
    e.parameter.action
      ? e.parameter.action
      : "health";

  try {

    let result;

    switch (action) {

      case "health":
        result = apiHealth_();
        break;

      case "me":
        result = apiMe_();
        break;

      case "dashboard":
        result = apiDashboard_();
        break;

      case "certificates":
        result = apiCertificates_(
          e.parameter.status || "",
          e.parameter.department || ""
        );
        break;

      case "certificate360":
        result = apiCertificate360_(
          e.parameter.id || ""
        );
        break;

      case "employees":
        result = apiEmployees_();
        break;

      case "employee360":
        result = apiEmployee360_(
          e.parameter.id || ""
        );
        break;

      case "renewals":
        result = apiRenewals_();
        break;

      case "documents":
        result = apiDocuments_(
          e.parameter.certificateId || ""
        );
        break;

      case "reports":
        result = apiReports_();
        break;
        
      case 'documentDownload':
         return jsonResponse_(authorizeDocumentDownload(e.parameter.id)
        );
      default:
        result = {
          success: false,
          error: "UNKNOWN_ACTION"
        };
    }

    return jsonResponse_(result);

  } catch (error) {

    logAudit_(
      getCurrentUserEmail_(),
      "API_ERROR",
      "SYSTEM",
      "",
      error.message,
      "ERROR"
    );

    return jsonResponse_({
      success: false,
      error: "SERVER_ERROR",
      message: error.message
    });
  }
}


/************************************************************
 * POST API
 ************************************************************/

function doPost(e) {
  try {
    const body=e&&e.postData&&e.postData.contents?JSON.parse(e.postData.contents):{};
    const action=String(body.action||"").trim();
    if(action==="passwordLogin") return jsonResponse_(passwordLogin_(body.email,body.password));
    if(action==="ccmExecute"){const session=validateCcmSession_(body.sessionToken);if(!session.valid)return jsonResponse_({success:false,error:"SESSION_EXPIRED",message:session.message});CCM_REQUEST_USER_EMAIL=session.email;return jsonResponse_(ccmExecute(body.targetAction||"",body.params||{}));}
    return jsonResponse_({success:false,error:"UNKNOWN_POST_ACTION"});
  } catch(error){return jsonResponse_({success:false,error:"POST_ERROR",message:error.message});}
}

/************************************************************
 * HEALTH CHECK
 ************************************************************/

function apiHealth_() {

  return {
    success: true,
    system: CONFIG.SYSTEM,
    company: CONFIG.COMPANY,
    status: "ONLINE",
    timestamp: new Date().toISOString(),
    version: "1.0.0"
  };
}


/************************************************************
 * EMAIL/PASSWORD AUTHENTICATION
 ************************************************************/
function hashCcmPassword_(password){const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(password||""),Utilities.Charset.UTF_8);return bytes.map(b=>{const v=b<0?b+256:b;return(v<16?"0":"")+v.toString(16)}).join("");}
function passwordLogin_(email,password){const e=String(email||"").trim().toLowerCase();const p=String(password||"");if(!e||!p)return{success:false,error:"LOGIN_REQUIRED",message:"Email and password are required."};const u=findUserByEmail_(e);if(!u)return{success:false,error:"INVALID_CREDENTIALS",message:"Invalid email or password."};if(String(u.Status||"").toUpperCase()!=="ACTIVE")return{success:false,error:"USER_INACTIVE",message:"User account is not active."};const stored=String(u.Password_Hash||"").trim().toLowerCase();if(!stored)return{success:false,error:"PASSWORD_NOT_CONFIGURED",message:"Password is not configured for this user. Ask CCM Admin to set it."};if(stored!==hashCcmPassword_(p))return{success:false,error:"INVALID_CREDENTIALS",message:"Invalid email or password."};const token=Utilities.getUuid()+Utilities.getUuid().replace(/-/g,"");CacheService.getScriptCache().put("CCM_SESSION_"+token,e,21600);updateLastLogin_(u.User_ID);logAudit_(e,"LOGIN","AUTH",u.User_ID||"","Successful email/password portal login","SUCCESS");return{success:true,authenticated:true,authorized:true,sessionToken:token,user:sanitizeUser_(u),permissions:getUserPermissions_(u.Role_ID)};}
function validateCcmSession_(token){const t=String(token||"").trim();if(!t)return{valid:false,message:"Session token is missing."};const e=CacheService.getScriptCache().get("CCM_SESSION_"+t);if(!e)return{valid:false,message:"Session expired. Please sign in again."};const u=findUserByEmail_(e);if(!u||String(u.Status||"").toUpperCase()!=="ACTIVE")return{valid:false,message:"User account is inactive or unavailable."};return{valid:true,email:e};}
function setCCMUserPassword(email, password){const e=String(email||"").trim().toLowerCase();if(!password||String(password).length<8)throw new Error("Password must be at least 8 characters.");const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.USERS);if(!sh)throw new Error("Sheet not found: "+CONFIG.SHEETS.USERS);const vals=sh.getDataRange().getValues();const h=vals[0].map(String);const ei=h.indexOf("Email");if(ei<0)throw new Error("06_USERS is missing Email column.");let pi=h.indexOf("Password_Hash");if(pi<0){pi=h.length;sh.getRange(1,pi+1).setValue("Password_Hash");}for(let i=1;i<vals.length;i++){if(String(vals[i][ei]).trim().toLowerCase()===e){sh.getRange(i+1,pi+1).setValue(hashCcmPassword_(password));return{success:true,email:e,message:"CCM password configured."};}}throw new Error("User not found: "+e);}

/************************************************************
 * CURRENT USER
 ************************************************************/

function apiMe_() {

  const email = getCurrentUserEmail_();

  if (!email) {

    return {
      success: false,
      authenticated: false,
      message: "Unable to identify current user."
    };
  }

  const user = findUserByEmail_(email);

  if (!user) {

    logAudit_(
      email,
      "LOGIN",
      "AUTH",
      "",
      "User not registered",
      "DENIED"
    );

    return {
      success: false,
      authenticated: true,
      authorized: false,
      email: email,
      message: "User is not registered in 06_USERS."
    };
  }

  if (String(user.Status).toUpperCase() !== "ACTIVE") {

    return {
      success: false,
      authenticated: true,
      authorized: false,
      email: email,
      message: "User account is not active."
    };
  }

  updateLastLogin_(user.User_ID);

  logAudit_(
    email,
    "LOGIN",
    "AUTH",
    user.User_ID,
    "Successful portal login",
    "SUCCESS"
  );

  return {
    success: true,
    authenticated: true,
    authorized: true,
    user: sanitizeUser_(user),
    permissions: getUserPermissions_(user.Role_ID)
  };
}


/************************************************************
 * DASHBOARD
 ************************************************************/

function apiDashboard_() {
  requirePermission_('PERM-DASHBOARD-VIEW');

  const rows = getSheetObjects_('01_CERTIFICATIONS');

  const now = new Date();
  const today = new Date(
    now.getFullYear(),
    now.getMonth(),
    now.getDate()
  );

  let total = rows.length;
  let active = 0;
  let expired = 0;
  let renewal = 0;
  let mandatory = 0;
  let compliant = 0;

  const expiryRisk = {
    '0-7': 0,
    '8-15': 0,
    '16-30': 0,
    '31-60': 0,
    '61-90': 0,
    '90+': 0,
    'EXPIRED': 0
  };

  const departments = {};

  rows.forEach(function (r) {

    let expiry = null;

    if (r.Expiry_Date) {
      expiry = new Date(r.Expiry_Date);

      if (isNaN(expiry.getTime())) {
        expiry = null;
      }
    }

    let daysRemaining = null;

    if (expiry) {
      const expiryDay = new Date(
        expiry.getFullYear(),
        expiry.getMonth(),
        expiry.getDate()
      );

      daysRemaining = Math.floor(
        (expiryDay.getTime() - today.getTime()) /
        (1000 * 60 * 60 * 24)
      );
    }

    const status = String(r.Status || '')
      .trim()
      .toUpperCase();

    const mandatoryFlag = String(r.Mandatory || '')
      .trim()
      .toUpperCase();

    /*
     * EXPIRED
     */
    const isExpired =
      expiry &&
      daysRemaining < 0;

    /*
     * ACTIVE
     *
     * An expired certificate can NEVER be active.
     */
    const isActive =
      !isExpired &&
      status === 'ACTIVE';

    /*
     * RENEWAL DUE
     *
     * Any certificate expiring within 90 days.
     */
    const isRenewalDue =
      expiry &&
      daysRemaining >= 0 &&
      daysRemaining <= 90;

    /*
     * KPI COUNTS
     */

    if (isExpired) {
      expired++;
    }

    if (isActive) {
      active++;
    }

    if (isRenewalDue) {
      renewal++;
    }

    if (
      mandatoryFlag === 'YES' ||
      mandatoryFlag === 'TRUE'
    ) {
      mandatory++;
    }

    /*
     * COMPLIANCE
     *
     * Valid active certificates only.
     */
    if (isActive && expiry) {
      compliant++;
    }

    /*
     * EXPIRY RISK
     */

    if (expiry) {

      if (daysRemaining < 0) {
        expiryRisk['EXPIRED']++;

      } else if (daysRemaining <= 7) {
        expiryRisk['0-7']++;

      } else if (daysRemaining <= 15) {
        expiryRisk['8-15']++;

      } else if (daysRemaining <= 30) {
        expiryRisk['16-30']++;

      } else if (daysRemaining <= 60) {
        expiryRisk['31-60']++;

      } else if (daysRemaining <= 90) {
        expiryRisk['61-90']++;

      } else {
        expiryRisk['90+']++;
      }
    }

    /*
     * DEPARTMENT STATISTICS
     */

    const department =
      String(r.Department || 'Unassigned').trim() ||
      'Unassigned';

    if (!departments[department]) {
      departments[department] = {
        total: 0,
        compliant: 0,
        compliance: 0
      };
    }

    departments[department].total++;

    if (isActive && expiry) {
      departments[department].compliant++;
    }
  });

  /*
   * DEPARTMENT COMPLIANCE %
   */

  Object.keys(departments).forEach(function (department) {

    const d = departments[department];

    d.compliance = d.total > 0
      ? Math.round(
          (d.compliant / d.total) * 100
        )
      : 0;
  });

  /*
   * OVERALL COMPLIANCE %
   */

  const compliance = total > 0
    ? Math.round(
        (compliant / total) * 100
      )
    : 0;

  /*
   * RECENT RENEWALS
   */

  const recentRenewals = rows
    .filter(function (r) {

      return String(r.Renewal_Status || '')
        .trim()
        .toUpperCase() === 'RENEWED';

    })
    .slice(-10)
    .reverse()
    .map(function (r) {

      return {
        Certificate_ID: r.Certificate_ID || '',
        Employee_ID: r.Employee_ID || '',
        Employee_Name: r.Employee_Name || '',
        Certification_Name: r.Certification_Name || '',
        Updated_Date: r.Updated_Date || ''
      };

    });

  /*
   * FINAL RESPONSE
   */

  const authorizationKpi = getAuthorizationDashboardKpis_();

  return {
    success: true,

    authorizationKpi: authorizationKpi,

    kpi: {
      total: total,
      active: active,
      expired: expired,
      renewal: renewal,
      mandatory: mandatory,
      compliance: compliance
    },

    expiryRisk: expiryRisk,

    departments: departments,

    recentRenewals: recentRenewals,

    generatedAt: new Date().toISOString()
  };
}


/************************************************************
 * CERTIFICATES
 ************************************************************/

function apiCertificates_(status, department) {

  const auth = requirePermission_("PERM-CERT-VIEW");

  if (!auth.allowed) {

    /*
     * Employee fallback:
     * employees can see their own records.
     */

    const own =
      requirePermission_("PERM-CERT-OWN-VIEW");

    if (!own.allowed) {
      return own.response;
    }

    const email = getCurrentUserEmail_();
    const user = findUserByEmail_(email);

    const employeeId = user
      ? user.Employee_ID
      : "";

    const records =
      getSheetObjects_(
        CONFIG.SHEETS.CERTIFICATIONS
      )
      .filter(
        x => String(x.Employee_ID) ===
             String(employeeId)
      );

    return {
      success: true,
      count: records.length,
      data: records.map(sanitizeCertificate_)
    };
  }

  let records =
    getSheetObjects_(
      CONFIG.SHEETS.CERTIFICATIONS
    );

  if (status) {

    records = records.filter(
      x =>
        String(x.Status).toUpperCase() ===
        String(status).toUpperCase()
    );
  }

  if (department) {

    records = records.filter(
      x =>
        String(x.Department).toLowerCase() ===
        String(department).toLowerCase()
    );
  }

  return {
    success: true,
    count: records.length,
    data: records.map(sanitizeCertificate_)
  };
}


/************************************************************
 * CERTIFICATE 360
 ************************************************************/

function apiCertificate360_(certificateId) {

  if (!certificateId) {

    return {
      success: false,
      error: "CERTIFICATE_ID_REQUIRED"
    };
  }

  const certificate =
    findByField_(
      CONFIG.SHEETS.CERTIFICATIONS,
      "Certificate_ID",
      certificateId
    );

  if (!certificate) {

    return {
      success: false,
      error: "CERTIFICATE_NOT_FOUND"
    };
  }

  const access =
    canViewCertificate_(certificate);

  if (!access.allowed) {
    return access.response;
  }

  const documents =
    getSheetObjects_(
      CONFIG.SHEETS.DOCUMENTS
    )
    .filter(
      d =>
        String(d.Certificate_ID) ===
        String(certificateId)
    )
    .map(d =>
      sanitizeDocumentForUser_(d)
    );

  const employee =
    findByField_(
      CONFIG.SHEETS.EMPLOYEES,
      "Employee_ID",
      certificate.Employee_ID
    );

  const days =
    daysRemaining_(certificate.Expiry_Date);

  return {
    success: true,

    certificate: {
      ...sanitizeCertificate_(certificate),
      daysRemaining: days,
      calculatedRenewalStatus:
        calculateRenewalStatus_(days)
    },

    employee:
      employee
        ? sanitizeEmployee_(employee)
        : null,

    documents: documents,

    notifications:
      getNotificationHistory_(
        certificate.Certificate_ID
      ),

    audit:
      getAuditHistory_(
        certificate.Certificate_ID
      )
  };
}


/************************************************************
 * EMPLOYEES
 ************************************************************/

function apiEmployees_() {

  const auth =
    requirePermission_("PERM-EMPLOYEE-VIEW");

  if (!auth.allowed) {
    return auth.response;
  }

  const employees =
    getSheetObjects_(
      CONFIG.SHEETS.EMPLOYEES
    );

  return {
    success: true,
    count: employees.length,
    data: employees.map(sanitizeEmployee_)
  };
}


/************************************************************
 * EMPLOYEE 360
 ************************************************************/

function apiEmployee360_(employeeId) {

  if (!employeeId) {

    return {
      success: false,
      error: "EMPLOYEE_ID_REQUIRED"
    };
  }

  const employee =
    findByField_(
      CONFIG.SHEETS.EMPLOYEES,
      "Employee_ID",
      employeeId
    );

  if (!employee) {

    return {
      success: false,
      error: "EMPLOYEE_NOT_FOUND"
    };
  }

  const email =
    getCurrentUserEmail_();

  const currentUser =
    findUserByEmail_(email);

  const isOwn =
    currentUser &&
    String(currentUser.Employee_ID) ===
    String(employeeId);

  if (!isOwn) {

    const auth =
      requirePermission_("PERM-EMPLOYEE-VIEW");

    if (!auth.allowed) {
      return auth.response;
    }
  }

  const certifications =
    getSheetObjects_(
      CONFIG.SHEETS.CERTIFICATIONS
    )
    .filter(
      x =>
        String(x.Employee_ID) ===
        String(employeeId)
    );

  return {
    success: true,

    employee:
      sanitizeEmployee_(employee),

    certifications:
      certifications.map(sanitizeCertificate_),

    summary:
      employeeCertificationSummary_(
        certifications
      )
  };
}


/************************************************************
 * RENEWALS
 ************************************************************/

function apiRenewals_() {

  const auth =
    requirePermission_("PERM-RENEWAL-VIEW");

  if (!auth.allowed) {
    return auth.response;
  }

  const certificates =
    getSheetObjects_(
      CONFIG.SHEETS.CERTIFICATIONS
    );

  const output = [];

  certificates.forEach(cert => {

    const days =
      daysRemaining_(cert.Expiry_Date);

    if (
      days !== null &&
      days <= 90
    ) {

      output.push({

        certificateId:
          cert.Certificate_ID,

        employeeId:
          cert.Employee_ID,

        employee:
          cert.Employee_Name,

        department:
          cert.Department,

        certification:
          cert.Certification_Name,

        expiryDate:
          formatDate_(cert.Expiry_Date),

        daysRemaining:
          days,

        priority:
          renewalPriority_(days),

        status:
          calculateRenewalStatus_(days)
      });
    }
  });

  output.sort(
    (a, b) =>
      a.daysRemaining -
      b.daysRemaining
  );

  return {
    success: true,
    count: output.length,
    data: output
  };
}


/************************************************************
 * DOCUMENTS
 ************************************************************/

function apiDocuments_(certificateId) {

  const records =
    getSheetObjects_(
      CONFIG.SHEETS.DOCUMENTS
    );

  let filtered = records;

  if (certificateId) {

    filtered =
      records.filter(
        d =>
          String(d.Certificate_ID) ===
          String(certificateId)
      );
  }

  const output = [];

  filtered.forEach(document => {

    const access =
      canAccessDocument_(document);

    if (access.allowed) {

      output.push(
        sanitizeDocumentForUser_(document)
      );
    }
  });

  return {
    success: true,
    count: output.length,
    data: output
  };
}


/************************************************************
 * DOCUMENT DOWNLOAD AUTHORIZATION
 *
 * IMPORTANT:
 * The frontend MUST NOT decide whether a document
 * is downloadable.
 *
 * This function performs the actual authorization.
 ************************************************************/

function authorizeDocumentDownload(
  documentId
) {

  const document =
    findByField_(
      CONFIG.SHEETS.DOCUMENTS,
      "Document_ID",
      documentId
    );

  if (!document) {

    return {
      success: false,
      allowed: false,
      error: "DOCUMENT_NOT_FOUND"
    };
  }

  const access =
    canAccessDocument_(document);

  logAudit_(
    getCurrentUserEmail_(),
    "DOCUMENT_DOWNLOAD",
    "DOCUMENT",
    documentId,
    document.Document_Name || "",
    access.allowed
      ? "SUCCESS"
      : "DENIED"
  );

  if (!access.allowed) {

    return {
      success: false,
      allowed: false,
      error: "ACCESS_DENIED"
    };
  }

  /*
   * PUBLIC documents can return their public URL.
   */

  if (
    String(document.Classification)
      .toUpperCase() === "PUBLIC"
  ) {

    return {
      success: true,
      allowed: true,
      type: "PUBLIC",
      url: document.Public_URL || ""
    };
  }

  /*
   * PRIVATE document.
   *
   * We do NOT expose Drive URL directly unless
   * authorization succeeds.
   */

  if (!document.Drive_File_ID) {

    return {
      success: false,
      allowed: false,
      error: "DRIVE_FILE_NOT_CONFIGURED"
    };
  }

  try {

    const file =
      DriveApp.getFileById(
        document.Drive_File_ID
      );

    /*
     * Return only after authorization.
     */

    return {
      success: true,
      allowed: true,
      type: "PRIVATE",
      fileName: file.getName(),
      url: file.getUrl()
    };

  } catch (error) {

    return {
      success: false,
      allowed: false,
      error: "DRIVE_ACCESS_ERROR"
    };
  }
}


/************************************************************
 * CREATE CERTIFICATE
 ************************************************************/

function createCertificate_(data) {

  const auth =
    requirePermission_("PERM-CERT-CREATE");

  if (!auth.allowed) {
    return auth.response;
  }

  if (!data) {

    return {
      success: false,
      error: "DATA_REQUIRED"
    };
  }

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        CONFIG.SHEETS.CERTIFICATIONS
      );

  const certificateId =
    generateId_("CERT");

  const now =
    new Date();

  const row = [

    certificateId,

    data.Employee_ID || "",

    data.Employee_Name || "",

    data.Department || "",

    data.Designation || "",

    data.Certification_Name || "",

    data.Certification_Category || "",

    data.Issuing_Body || "",

    data.Certificate_Number || "",

    data.Issue_Date || "",

    data.Expiry_Date || "",

    data.Mandatory || "NO",

    "ACTIVE",

    calculateRenewalStatus_(
      daysRemaining_(data.Expiry_Date)
    ),

    data.Document_ID || "",

    data.Remarks || "",

    now,

    now
  ];

  sheet
    .getRange(
      sheet.getLastRow() + 1,
      1,
      1,
      row.length
    )
    .setValues([row]);

  logAudit_(
    getCurrentUserEmail_(),
    "CREATE",
    "CERTIFICATE",
    certificateId,
    "Certificate created",
    "SUCCESS"
  );

  return {
    success: true,
    certificateId: certificateId
  };
}


/************************************************************
 * UPDATE CERTIFICATE
 ************************************************************/

function updateCertificate_(
  certificateId,
  data
) {

  const auth =
    requirePermission_("PERM-CERT-EDIT");

  if (!auth.allowed) {
    return auth.response;
  }

  if (!certificateId) {

    return {
      success: false,
      error: "CERTIFICATE_ID_REQUIRED"
    };
  }

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        CONFIG.SHEETS.CERTIFICATIONS
      );

  const values =
    sheet.getDataRange().getValues();

  const headers =
    values[0];

  const idIndex =
    headers.indexOf("Certificate_ID");

  const rowIndex =
    values.findIndex(
      (row, index) =>
        index > 0 &&
        String(row[idIndex]) ===
        String(certificateId)
    );

  if (rowIndex === -1) {

    return {
      success: false,
      error: "CERTIFICATE_NOT_FOUND"
    };
  }

  const row =
    values[rowIndex];

  Object.keys(data || {}).forEach(key => {

    const index =
      headers.indexOf(key);

    if (index !== -1) {
      row[index] = data[key];
    }
  });

  const expiryIndex =
    headers.indexOf("Expiry_Date");

  const renewalIndex =
    headers.indexOf("Renewal_Status");

  if (expiryIndex !== -1 &&
      renewalIndex !== -1) {

    row[renewalIndex] =
      calculateRenewalStatus_(
        daysRemaining_(
          row[expiryIndex]
        )
      );
  }

  row[
    headers.indexOf("Updated_Date")
  ] = new Date();

  sheet
    .getRange(
      rowIndex + 1,
      1,
      1,
      row.length
    )
    .setValues([row]);

  logAudit_(
    getCurrentUserEmail_(),
    "UPDATE",
    "CERTIFICATE",
    certificateId,
    "Certificate updated",
    "SUCCESS"
  );

  return {
    success: true,
    certificateId: certificateId
  };
}


/************************************************************
 * RENEWAL ENGINE
 ************************************************************/

function runRenewalCheck() {

  const certificates =
    getSheetObjects_(
      CONFIG.SHEETS.CERTIFICATIONS
    );

  const rules =
    getSheetObjects_(
      CONFIG.SHEETS.RENEWAL_RULES
    );

  let sent = 0;
  let skipped = 0;
  let expired = 0;

  certificates.forEach(cert => {

    const days =
      daysRemaining_(cert.Expiry_Date);

    if (days === null) {
      return;
    }

    if (days < 0) {
      expired++;
    }

    const rule =
      findRenewalRule_(
        rules,
        days
      );

    if (!rule) {
      return;
    }

    const employee =
      findByField_(
        CONFIG.SHEETS.EMPLOYEES,
        "Employee_ID",
        cert.Employee_ID
      );

    if (!employee) {
      return;
    }

    const recipients =
      buildRecipients_(
        rule,
        employee
      );

    const notificationType =
      rule.Email_Type;

    /*
     * Duplicate prevention.
     */

    if (
      notificationAlreadySent_(
        cert.Certificate_ID,
        notificationType
      )
    ) {

      skipped++;
      return;
    }

    try {

      sendRenewalEmail_(
        cert,
        employee,
        rule,
        recipients,
        days
      );

      writeNotificationLog_(
        cert,
        notificationType,
        days,
        recipients,
        "SENT"
      );

      sent++;

    } catch (error) {

      writeNotificationLog_(
        cert,
        notificationType,
        days,
        recipients,
        "FAILED",
        error.message
      );
    }

  });

  return {
    success: true,
    checked: certificates.length,
    sent: sent,
    skipped: skipped,
    expired: expired,
    timestamp: new Date().toISOString()
  };
}


/************************************************************
 * SEND RENEWAL EMAIL
 ************************************************************/

function sendRenewalEmail_(
  cert,
  employee,
  rule,
  recipients,
  days
) {

  if (!recipients.to) {

    throw new Error(
      "No recipient email available."
    );
  }

  const priority =
    renewalPriority_(days);

  const subject =
    days < 0
      ? "ACTION REQUIRED - Certification Expired - " +
        cert.Certification_Name
      : days === 0
        ? "URGENT - Certification Expires Today - " +
          cert.Certification_Name
        : "Certification Renewal Reminder - " +
          cert.Certification_Name +
          " - " +
          days +
          " Days Remaining";

  const statusText =
    days < 0
      ? "EXPIRED"
      : days === 0
        ? "EXPIRES TODAY"
        : days + " DAYS REMAINING";

  const htmlBody = `

  <div style="
    font-family:Arial,sans-serif;
    max-width:700px;
    margin:auto;
    border:1px solid #ddd;
    border-radius:8px;
    overflow:hidden;
  ">

    <div style="
      background:#17365D;
      color:white;
      padding:20px;
    ">

      <h2 style="margin:0;">
        ${CONFIG.SYSTEM}
      </h2>

      <div style="
        margin-top:5px;
        opacity:.85;
      ">
        ${CONFIG.COMPANY}
      </div>

    </div>

    <div style="padding:25px;">

      <h3>
        Certification Renewal Notification
      </h3>

      <p>
        Dear ${employee.Employee_Name || "Employee"},
      </p>

      <p>
        This is an automated certification
        compliance notification.
      </p>

      <table style="
        border-collapse:collapse;
        width:100%;
      ">

        <tr>
          <td style="padding:8px;border:1px solid #ddd;">
            Certification
          </td>

          <td style="padding:8px;border:1px solid #ddd;">
            <b>${cert.Certification_Name}</b>
          </td>
        </tr>

        <tr>
          <td style="padding:8px;border:1px solid #ddd;">
            Employee
          </td>

          <td style="padding:8px;border:1px solid #ddd;">
            ${cert.Employee_Name}
          </td>
        </tr>

        <tr>
          <td style="padding:8px;border:1px solid #ddd;">
            Certificate Number
          </td>

          <td style="padding:8px;border:1px solid #ddd;">
            ${cert.Certificate_Number || "-"}
          </td>
        </tr>

        <tr>
          <td style="padding:8px;border:1px solid #ddd;">
            Expiry Date
          </td>

          <td style="padding:8px;border:1px solid #ddd;">
            ${formatDate_(cert.Expiry_Date)}
          </td>
        </tr>

        <tr>
          <td style="padding:8px;border:1px solid #ddd;">
            Status
          </td>

          <td style="
            padding:8px;
            border:1px solid #ddd;
            font-weight:bold;
          ">
            ${statusText}
          </td>
        </tr>

      </table>

      <p style="
        margin-top:20px;
        padding:15px;
        background:#f5f7fa;
        border-radius:6px;
      ">

        Priority:
        <b>${priority}</b>

      </p>

      <p>
        Please complete the required renewal
        action and submit the renewed certificate
        to the appropriate department.
      </p>

      <hr>

      <p style="
        font-size:12px;
        color:#777;
      ">
        This is an automated message generated
        by the Central Certification Manager.
        Please do not reply directly to this email.
      </p>

    </div>

  </div>

  `;

  MailApp.sendEmail({
    to: recipients.to,
    cc: recipients.cc,
    subject: subject,
    htmlBody: htmlBody
  });
}


/************************************************************
 * NOTIFICATION LOG
 ************************************************************/

function notificationAlreadySent_(certificateId, notificationType) {
  const sheet = getSheet_(CONFIG.SHEETS.NOTIFICATION_LOG);
  if (!sheet) return false;

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return false;

  const headers = values[0].map(function(h) { return String(h).trim(); });
  const ci = headers.indexOf("Certificate_ID");
  const ti = headers.indexOf("Notification_Type");
  const si = headers.indexOf("Status");

  if (ci === -1 || ti === -1 || si === -1) return false;

  const targetCertificate = String(certificateId || "").trim();
  const targetType = String(notificationType || "").trim().toUpperCase();

  for (let i=1; i<values.length; i++) {
    if (
      String(values[i][ci] || "").trim() === targetCertificate &&
      String(values[i][ti] || "").trim().toUpperCase() === targetType &&
      String(values[i][si] || "").trim().toUpperCase() === "SENT"
    ) return true;
  }
  return false;
}


function writeNotificationLog_(
  cert,
  type,
  days,
  recipients,
  status,
  errorMessage
) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        CONFIG.SHEETS.NOTIFICATION_LOG
      );

  const row = [

    generateId_("NOT"),

    cert.Certificate_ID,

    cert.Employee_ID,

    type,

    days,

    recipients.to,

    recipients.cc,

    new Date(),

    status,

    errorMessage || ""
  ];

  sheet
    .getRange(
      sheet.getLastRow() + 1,
      1,
      1,
      row.length
    )
    .setValues([row]);
}


/************************************************************
 * RECIPIENT BUILDER
 ************************************************************/

function buildRecipients_(
  rule,
  employee
) {

  let to = [];
  let cc = [];

  const recipients =
    String(rule.Recipient || "")
      .split(";")
      .map(x => x.trim())
      .filter(Boolean);

  const ccRecipients =
    String(rule.CC_Recipient || "")
      .split(";")
      .map(x => x.trim())
      .filter(Boolean);

  recipients.forEach(type => {

    if (
      type.toUpperCase() === "EMPLOYEE" &&
      employee.Email
    ) {
      to.push(employee.Email);
    }

    if (
      type.toUpperCase() === "HR" &&
      employee.HR_Email
    ) {
      to.push(employee.HR_Email);
    }

    if (
      type.toUpperCase() === "MANAGER" &&
      employee.Manager_Email
    ) {
      to.push(employee.Manager_Email);
    }

  });

  ccRecipients.forEach(type => {

    if (
      type.toUpperCase() === "EMPLOYEE" &&
      employee.Email
    ) {
      cc.push(employee.Email);
    }

    if (
      type.toUpperCase() === "HR" &&
      employee.HR_Email
    ) {
      cc.push(employee.HR_Email);
    }

    if (
      type.toUpperCase() === "MANAGER" &&
      employee.Manager_Email
    ) {
      cc.push(employee.Manager_Email);
    }

  });

  /*
   * Remove duplicates.
   */

  to = [...new Set(to)];
  cc = [...new Set(cc)];

  /*
   * Don't CC someone already in TO.
   */

  cc =
    cc.filter(
      email => !to.includes(email)
    );

  return {
    to: to.join(","),
    cc: cc.join(",")
  };
}


/************************************************************
 * RENEWAL RULE MATCH
 ************************************************************/

function findRenewalRule_(rules, days) {
  const d = Number(days);
  if (!isFinite(d)) return null;

  const enabledRules = (rules || []).filter(function(r) {
    return String(r.Enabled || "").trim().toUpperCase() === "TRUE";
  });

  if (d < 0) {
    return enabledRules.find(function(r) {
      return String(r.Email_Type || "").trim().toUpperCase() === "EXPIRED" ||
             Number(r.Days_Before_Expiry) === -1;
    }) || {
      Rule_ID:"RULE-EXP", Days_Before_Expiry:-1, Email_Type:"EXPIRED",
      Recipient:"HR", CC_Recipient:"MANAGER;COMPLIANCE",
      Enabled:"TRUE", Priority:"CRITICAL"
    };
  }

  if (d === 0) {
    return enabledRules.find(function(r) {
      return Number(r.Days_Before_Expiry) === 0;
    }) || {
      Rule_ID:"RULE-000", Days_Before_Expiry:0, Email_Type:"EXPIRES_TODAY",
      Recipient:"HR", CC_Recipient:"MANAGER",
      Enabled:"TRUE", Priority:"CRITICAL"
    };
  }

  const reached = enabledRules
    .map(function(r) {
      return { rule:r, threshold:Number(r.Days_Before_Expiry) };
    })
    .filter(function(x) {
      return isFinite(x.threshold) && x.threshold > 0 && x.threshold >= d;
    })
    .sort(function(a,b) { return a.threshold - b.threshold; });

  return reached.length ? reached[0].rule : null;
}


/************************************************************
 * ACCESS CONTROL
 ************************************************************/

function requirePermission_(
  permissionId
) {

  const email =
    getCurrentUserEmail_();

  if (!email) {

    return {
      allowed: false,
      response: {
        success: false,
        authenticated: false,
        error: "AUTHENTICATION_REQUIRED"
      }
    };
  }

  const user =
    findUserByEmail_(email);

  if (!user) {

    logAudit_(
      email,
      "AUTHORIZATION",
      "RBAC",
      permissionId,
      "User not found",
      "DENIED"
    );

    return {
      allowed: false,
      response: {
        success: false,
        authorized: false,
        error: "USER_NOT_REGISTERED"
      }
    };
  }

  if (
    String(user.Status).toUpperCase() !==
    "ACTIVE"
  ) {

    return {
      allowed: false,
      response: {
        success: false,
        authorized: false,
        error: "USER_INACTIVE"
      }
    };
  }

  const permissions =
    getUserPermissions_(
      user.Role_ID
    );

  const allowed =
    permissions.includes(
      permissionId
    );

  if (!allowed) {

    logAudit_(
      email,
      "AUTHORIZATION",
      "RBAC",
      permissionId,
      "Permission denied",
      "DENIED"
    );

    return {
      allowed: false,
      response: {
        success: false,
        authorized: false,
        error: "PERMISSION_DENIED",
        permission: permissionId
      }
    };
  }

  return {
    allowed: true,
    user: user
  };
}


/************************************************************
 * USER PERMISSIONS
 ************************************************************/

function getUserPermissions_(
  roleId
) {

  const rows =
    getSheetObjects_(
      CONFIG.SHEETS.ROLE_PERMISSIONS
    );

  return rows
    .filter(
      r =>
        String(r.Role_ID) ===
        String(roleId) &&
        String(r.Allowed).toUpperCase() ===
        "TRUE"
    )
    .map(
      r => r.Permission_ID
    );
}


/************************************************************
 * CERTIFICATE VIEW ACCESS
 ************************************************************/

function canViewCertificate_(
  certificate
) {

  const email =
    getCurrentUserEmail_();

  const user =
    findUserByEmail_(email);

  if (!user) {

    return {
      allowed: false,
      response: {
        success: false,
        error: "AUTHENTICATION_REQUIRED"
      }
    };
  }

  /*
   * Employee can see own record.
   */

  if (
    String(user.Employee_ID) ===
    String(certificate.Employee_ID)
  ) {

    const own =
      getUserPermissions_(
        user.Role_ID
      );

    if (
      own.includes(
        "PERM-CERT-OWN-VIEW"
      )
    ) {

      return {
        allowed: true
      };
    }
  }

  /*
   * Management / HR / Compliance.
   */

  const permission =
    getUserPermissions_(
      user.Role_ID
    );

  if (
    permission.includes(
      "PERM-CERT-VIEW"
    )
  ) {

    return {
      allowed: true
    };
  }

  return {
    allowed: false,
    response: {
      success: false,
      error: "CERTIFICATE_ACCESS_DENIED"
    }
  };
}


/************************************************************
 * DOCUMENT ACCESS
 ************************************************************/

function canAccessDocument_(
  document
) {

  const classification =
    String(
      document.Classification || "INTERNAL"
    ).toUpperCase();

  /*
   * PUBLIC
   */

  if (classification === "PUBLIC") {

    return {
      allowed: true
    };
  }

  /*
   * Everything else requires authenticated user.
   */

  const email =
    getCurrentUserEmail_();

  if (!email) {

    return {
      allowed: false,
      response: {
        success: false,
        error: "AUTHENTICATION_REQUIRED"
      }
    };
  }

  const user =
    findUserByEmail_(email);

  if (!user) {

    return {
      allowed: false
    };
  }

  if (
    String(user.Status).toUpperCase() !==
    "ACTIVE"
  ) {

    return {
      allowed: false
    };
  }

  /*
   * Admin
   */

  if (
    user.Role_ID ===
    "ROLE-ADMIN"
  ) {

    return {
      allowed: true
    };
  }

  /*
   * Employee private document
   */

  if (
    classification ===
    "EMPLOYEE_PRIVATE"
  ) {

    if (
      String(user.Employee_ID) ===
      String(
        getCertificateEmployeeId_(
          document.Certificate_ID
        )
      )
    ) {

      return {
        allowed:
          getUserPermissions_(
            user.Role_ID
          ).includes(
            "PERM-CERT-OWN-DOWNLOAD"
          )
      };
    }

    return {
      allowed: false
    };
  }

  /*
   * HR
   */

  if (
    user.Role_ID === "ROLE-HR"
  ) {

    return {
      allowed:
        classification !== "RESTRICTED"
    };
  }

  /*
   * Compliance
   */

  if (
    user.Role_ID ===
    "ROLE-COMPLIANCE"
  ) {

    return {
      allowed:
        classification === "INTERNAL" ||
        classification === "CONFIDENTIAL" ||
        classification === "RESTRICTED"
    };
  }

  /*
   * Manager
   */

  if (
    user.Role_ID === "ROLE-MANAGER"
  ) {

    return {
      allowed:
        classification === "INTERNAL"
    };
  }

  return {
    allowed: false
  };
}


/************************************************************
 * SANITIZE DOCUMENT
 ************************************************************/

function sanitizeDocumentForUser_(
  document
) {

  const classification =
    String(
      document.Classification || ""
    ).toUpperCase();

  const access =
    canAccessDocument_(
      document
    );

  return {

    Document_ID:
      document.Document_ID,

    Certificate_ID:
      document.Certificate_ID,

    Document_Name:
      document.Document_Name,

    Classification:
      classification,

    Login_Required:
      document.Login_Required,

    Download_Allowed:
      access.allowed,

    Status:
      document.Status,

    /*
     * Only public URL is exposed here.
     */

    Public_URL:
      classification === "PUBLIC"
        ? document.Public_URL || ""
        : ""
  };
}


/************************************************************
 * REPORTS
 ************************************************************/

function apiReports_() {

  const auth =
    requirePermission_(
      "PERM-REPORT-VIEW"
    );

  if (!auth.allowed) {
    return auth.response;
  }

  const certificates =
    getSheetObjects_(
      CONFIG.SHEETS.CERTIFICATIONS
    );

  const employees =
    getSheetObjects_(
      CONFIG.SHEETS.EMPLOYEES
    );

  const documents =
    getSheetObjects_(
      CONFIG.SHEETS.DOCUMENTS
    );

  const notifications =
    getSheetObjects_(
      CONFIG.SHEETS.NOTIFICATION_LOG
    );

  const audit =
    getSheetObjects_(
      CONFIG.SHEETS.AUDIT_LOG
    );

  const byCategory = {};
  const byIssuer = {};
  const byDepartment = {};
  const byStatus = {};

  let active = 0;
  let expired = 0;
  let expiring60 = 0;
  let expiring90 = 0;
  let mandatory = 0;

  certificates.forEach(cert => {

    const category =
      cert.Certification_Category ||
      "Unclassified";

    const issuer =
      cert.Issuing_Body ||
      "Unknown";

    const department =
      cert.Department ||
      "Unknown";

    const status =
      String(
        cert.Status ||
        ""
      ).toUpperCase();

    byCategory[category] =
      (byCategory[category] || 0) + 1;

    byIssuer[issuer] =
      (byIssuer[issuer] || 0) + 1;

    byDepartment[department] =
      (byDepartment[department] || 0) + 1;

    byStatus[status || "UNSPECIFIED"] =
      (byStatus[status || "UNSPECIFIED"] || 0) + 1;

    if (status === "ACTIVE") {
      active++;
    }

    const days =
      daysRemaining_(
        cert.Expiry_Date
      );

    if (days !== null) {

      if (days < 0) {
        expired++;
      }

      if (days >= 0 && days <= 60) {
        expiring60++;
      }

      if (days >= 0 && days <= 90) {
        expiring90++;
      }
    }

    const mandatoryValue =
      String(
        cert.Mandatory ||
        cert.Is_Mandatory ||
        ""
      ).toUpperCase();

    if (
      mandatoryValue === "TRUE" ||
      mandatoryValue === "YES"
    ) {
      mandatory++;
    }
  });

  const compliance =
    certificates.length > 0
      ? Math.round(
          (
            certificates.filter(cert => {
              const days =
                daysRemaining_(
                  cert.Expiry_Date
                );

              return (
                String(cert.Status || "")
                  .toUpperCase() === "ACTIVE" &&
                days !== null &&
                days >= 0
              );
            }).length /
            certificates.length
          ) * 100
        )
      : 0;

  return {
    success: true,
    total: certificates.length,
    active: active,
    expired: expired,
    expiring60: expiring60,
    expiring90: expiring90,
    mandatory: mandatory,
    compliance: compliance,
    employeeCount: employees.length,
    documentCount: documents.length,
    notificationCount: notifications.length,
    auditCount: audit.length,
    byCategory: byCategory,
    byIssuer: byIssuer,
    byDepartment: byDepartment,
    byStatus: byStatus,
    generatedAt:
      new Date().toISOString(),

    reportTypes: [
      ["CERTIFICATION_MASTER", "Certification Master Register"],
      ["CERTIFICATION_STATUS", "Certification Status Report"],
      ["EXPIRY_RENEWAL", "Expiry & Renewal Report"],
      ["DEPARTMENT_COMPLIANCE", "Department Compliance Report"],
      ["CERTIFICATION_TYPE", "Certification Type Report"],
      ["ISSUING_BODY", "Issuing Body Report"],
      ["DOCUMENT_CONTROL", "Document Control Report"],
      ["NOTIFICATION_HISTORY", "Notification History Report"],
      ["AUDIT_TRAIL", "Audit Trail Report"],
      ["CERTIFICATION_GAP", "Certification Gap Report"],
      ["IMPORT_EXPORT_AUTHORIZATION", "Import / Export Authorization Register"],
      ["MANAGEMENT_SUMMARY", "Management Summary"]
    ].map(function (item) {
      return {
        id: item[0],
        name: item[1]
      };
    })
  };
}


/************************************************************
 * CCM 12.0 — CONTROLLED REPORT GENERATION
 *
 * Generates XLSX + PDF in a private Drive folder.
 * No navigation/UI changes are required.
 ************************************************************/

function generateCCM12Report(
  reportType,
  options
) {

  const auth =
    requirePermission_(
      "PERM-REPORT-VIEW"
    );

  if (!auth.allowed) {
    return auth.response;
  }

  const normalizedType =
    String(
      reportType || ""
    ).trim().toUpperCase();

  if (!normalizedType) {
    throw new Error(
      "Report type is required."
    );
  }

  if (normalizedType === "AUDIT_TRAIL") {

    const auditAuth =
      requirePermission_(
        "PERM-AUDIT-VIEW"
      );

    if (!auditAuth.allowed) {
      return auditAuth.response;
    }
  }

  const report =
    buildCCM12ReportData_(
      normalizedType,
      options || {}
    );

  const reportId =
    generateId_("RPT");

  const generatedAt =
    new Date();

  const userEmail =
    getCurrentUserEmail_();

  /*
   * CCM 12.0 FIX:
   *
   * Do NOT call DriveApp here.
   * Do NOT create/find a Drive folder.
   * Do NOT attempt to enable the Advanced Drive API.
   *
   * SpreadsheetApp creates the controlled report workbook.
   * XLSX/PDF are exposed through the native Google Sheets
   * export endpoints using the Apps Script OAuth token.
   */

  const temp =
    SpreadsheetApp.create(
      CONFIG.COMPANY +
      " - " +
      report.title +
      " - " +
      reportId
    );

  try {

    const sheet =
      temp.getSheets()[0];

    sheet.setName("Report");
    try {
      const logoBlob=Utilities.newBlob(Utilities.base64Decode("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"),"image/png","gaerish-logistics-logo.png");
      const logoImage=sheet.insertImage(logoBlob,1,1);
      logoImage.setWidth(120);
      logoImage.setHeight(75);
    } catch(logoError) { console.warn("CCM report logo insertion failed: "+logoError.message); }

    const metadata = [
      ["Company", CONFIG.COMPANY],
      ["System", CONFIG.SYSTEM],
      ["Report ID", reportId],
      ["Report Name", report.title],
      ["Generated By", userEmail || ""],
      ["Generated At",
        Utilities.formatDate(
          generatedAt,
          CONFIG.TIMEZONE,
          "dd-MMM-yyyy HH:mm:ss"
        )
      ],
      ["Record Count", report.rows.length],
      ["Data Source", "CCM Google Sheets Master Data"]
    ];

    const width =
      Math.max(
        report.headers.length,
        2
      );

    const metadataRows =
      metadata.map(function (m) {

        const row = [];

        for (
          let i = 0;
          i < width;
          i++
        ) {

          row.push(
            i === 0
              ? m[0]
              : i === 1
                ? m[1]
                : ""
          );
        }

        return row;
      });

    const hasSections =
      Array.isArray(report.sections) &&
      report.sections.length > 0;

    if (!hasSections) {

      /*
       * Existing CCM report rendering path.
       * Kept unchanged for all existing report types.
       */
      const matrix =
        metadataRows
          .concat([
            Array(width).fill("")
          ])
          .concat([
            report.headers
              .slice(0, width)
          ])
          .concat(
            report.rows.map(function (row) {

              const out =
                row.slice(0, width);

              while (
                out.length < width
              ) {
                out.push("");
              }

              return out;
            })
          );

      sheet
        .getRange(
          1,
          1,
          matrix.length,
          width
        )
        .setValues(matrix);

      sheet
        .getRange(
          1,
          1,
          metadataRows.length,
          2
        )
        .setFontWeight("bold");

      const headerRow =
        metadataRows.length + 2;

      sheet
        .getRange(
          headerRow,
          1,
          1,
          width
        )
        .setFontWeight("bold");

      sheet.setFrozenRows(
        headerRow
      );

    } else {

      /*
       * Additive section rendering used only by reports that
       * explicitly provide report.sections.
       */
      let currentRow =
        metadataRows.length + 2;

      sheet
        .getRange(
          1,
          1,
          metadataRows.length,
          2
        )
        .setFontWeight("bold");

      function writeReportTable_(
        startRow,
        tableHeaders,
        tableRows,
        emptyMessage
      ) {
        const tableWidth =
          tableHeaders.length;

        const outputRows = [];

        outputRows.push(
          tableHeaders.slice()
        );

        if (
          tableRows &&
          tableRows.length
        ) {
          tableRows.forEach(function(row) {
            const out =
              row.slice(
                0,
                tableWidth
              );

            while (
              out.length < tableWidth
            ) {
              out.push("");
            }

            outputRows.push(out);
          });
        } else {
          const emptyRow =
            Array(tableWidth).fill("");

          emptyRow[0] =
            emptyMessage ||
            "No records available.";

          outputRows.push(emptyRow);
        }

        sheet
          .getRange(
            startRow,
            1,
            outputRows.length,
            tableWidth
          )
          .setValues(outputRows);

        sheet
          .getRange(
            startRow,
            1,
            1,
            tableWidth
          )
          .setFontWeight("bold");

        return (
          startRow +
          outputRows.length
        );
      }

      currentRow =
        writeReportTable_(
          currentRow,
          report.headers,
          report.rows,
          "No records available."
        );

      currentRow += 2;

      report.sections.forEach(function(section) {

        const sectionWidth =
          section.headers.length;

        /*
         * Section heading uses the existing report font/layout.
         * No new visual design is introduced.
         */
        sheet
          .getRange(
            currentRow,
            1,
            1,
            sectionWidth
          )
          .merge();

        sheet
          .getRange(
            currentRow,
            1
          )
          .setValue(
            section.title
          )
          .setFontWeight("bold");

        currentRow += 1;

        currentRow =
          writeReportTable_(
            currentRow,
            section.headers,
            section.rows,
            section.emptyMessage
          );

        currentRow += 2;
      });

      sheet.setFrozenRows(
        metadataRows.length + 2
      );
    }

    sheet
      .getDataRange()
      .setWrap(true);

    sheet.autoResizeColumns(
      1,
      width
    );

    /*
     * Keep the report workbook itself as the controlled
     * source document. No Drive service is required here.
     */

    SpreadsheetApp.flush();

    const spreadsheetId =
      temp.getId();

    const spreadsheetUrl =
      temp.getUrl();

    const token =
      ScriptApp.getOAuthToken();

    const xlsxUrl =
      "https://docs.google.com/spreadsheets/d/" +
      spreadsheetId +
      "/export?format=xlsx";

    const pdfUrl =
      "https://docs.google.com/spreadsheets/d/" +
      spreadsheetId +
      "/export?format=pdf" +
      "&size=A4" +
      "&portrait=false" +
      "&fitw=true" +
      "&sheetnames=false" +
      "&printtitle=false" +
      "&pagenumbers=true" +
      "&gridlines=false" +
      "&fzr=true";

    /*
     * Validate both export endpoints before returning them.
     */

    const xlsxResponse =
      UrlFetchApp.fetch(
        xlsxUrl,
        {
          headers: {
            Authorization:
              "Bearer " + token
          },
          muteHttpExceptions: true
        }
      );

    if (
      xlsxResponse.getResponseCode() !== 200
    ) {
      throw new Error(
        "XLSX export failed. HTTP " +
        xlsxResponse.getResponseCode()
      );
    }

    const pdfResponse =
      UrlFetchApp.fetch(
        pdfUrl,
        {
          headers: {
            Authorization:
              "Bearer " + token
          },
          muteHttpExceptions: true
        }
      );

    if (
      pdfResponse.getResponseCode() !== 200
    ) {
      throw new Error(
        "PDF export failed. HTTP " +
        pdfResponse.getResponseCode()
      );
    }

    const safeTitle =
      report.title.replace(
        /[\\/:*?"<>|]/g,
        "_"
      );

    const xlsxName =
      reportId +
      "_" +
      safeTitle +
      ".xlsx";

    const pdfName =
      reportId +
      "_" +
      safeTitle +
      ".pdf";

    const result = {
      success: true,

      reportId:
        reportId,

      reportType:
        normalizedType,

      title:
        report.title,

      generatedBy:
        userEmail || "",

      generatedAt:
        generatedAt.toISOString(),

      recordCount:
        report.rows.length,

      spreadsheet: {
        fileId:
          spreadsheetId,

        url:
          spreadsheetUrl,

        name:
          CONFIG.COMPANY +
          " - " +
          report.title +
          " - " +
          reportId
      },

      xlsx: {
        fileId:
          spreadsheetId,

        url:
          xlsxUrl,

        name:
          xlsxName
      },

      pdf: {
        fileId:
          spreadsheetId,

        url:
          pdfUrl,

        name:
          pdfName
      }
    };

    logAudit_(
      userEmail,
      "REPORT_GENERATED",
      "REPORTS",
      reportId,
      normalizedType +
      " generated; records=" +
      report.rows.length,
      "SUCCESS"
    );

    return result;

  } catch (error) {

    logAudit_(
      userEmail,
      "REPORT_GENERATED",
      "REPORTS",
      reportId,
      normalizedType +
      " failed: " +
      error.message,
      "ERROR"
    );

    throw error;
  }
}

/************************************************************
 * CCM 12.0 — REPORT DATA BUILDER
 ************************************************************/

function buildCCM12ReportData_(
  reportType,
  options
) {

  const certificates =
    getSheetObjects_(
      CONFIG.SHEETS.CERTIFICATIONS
    );

  const employees =
    getSheetObjects_(
      CONFIG.SHEETS.EMPLOYEES
    );

  const documents =
    getSheetObjects_(
      CONFIG.SHEETS.DOCUMENTS
    );

  const notifications =
    getSheetObjects_(
      CONFIG.SHEETS.NOTIFICATION_LOG
    );

  const audit =
    getSheetObjects_(
      CONFIG.SHEETS.AUDIT_LOG
    );

  const employeeMap = {};

  employees.forEach(e => {
    employeeMap[
      String(e.Employee_ID)
    ] = e;
  });

  const days = cert =>
    daysRemaining_(
      cert.Expiry_Date
    );

  const employeeName = cert =>
    cert.Employee_Name ||
    (
      employeeMap[
        String(cert.Employee_ID)
      ] || {}
    ).Employee_Name ||
    "";

  const certStatus = cert => {

    const d = days(cert);

    if (
      d !== null &&
      d < 0
    ) {
      return "EXPIRED";
    }

    const source =
      String(
        cert.Status || ""
      ).trim().toUpperCase();

    return source ||
      (
        d !== null
          ? "ACTIVE"
          : "UNKNOWN"
      );
  };

  let title = "";
  let headers = [];
  let rows = [];
  let sections = [];

  switch (reportType) {

    case "CERTIFICATION_MASTER":
      title =
        "Certification Master Register";
      headers = [
        "Certificate_ID",
        "Employee_ID",
        "Employee_Name",
        "Certification_Name",
        "Certification_Category",
        "Issuing_Body",
        "Department",
        "Issue_Date",
        "Expiry_Date",
        "Status",
        "Renewal_Status",
        "Mandatory"
      ];
      rows =
        certificates.map(c => [
          c.Certificate_ID || "",
          c.Employee_ID || "",
          employeeName(c),
          c.Certification_Name || "",
          c.Certification_Category || "",
          c.Issuing_Body || "",
          c.Department || "",
          formatDate_(c.Issue_Date),
          formatDate_(c.Expiry_Date),
          certStatus(c),
          c.Renewal_Status || "",
          c.Mandatory ||
            c.Is_Mandatory ||
            ""
        ]);
      break;

    case "CERTIFICATION_STATUS":
      title =
        "Certification Status Report";
      headers = [
        "Certificate_ID",
        "Employee_ID",
        "Employee_Name",
        "Certification_Name",
        "Department",
        "Expiry_Date",
        "Days_Remaining",
        "Status",
        "Renewal_Status"
      ];
      rows =
        certificates.map(c => [
          c.Certificate_ID || "",
          c.Employee_ID || "",
          employeeName(c),
          c.Certification_Name || "",
          c.Department || "",
          formatDate_(c.Expiry_Date),
          days(c) === null
            ? ""
            : days(c),
          certStatus(c),
          c.Renewal_Status || ""
        ]);
      break;

    case "EXPIRY_RENEWAL":
      title =
        "Expiry & Renewal Report";
      headers = [
        "Certificate_ID",
        "Employee_ID",
        "Employee_Name",
        "Certification_Name",
        "Department",
        "Expiry_Date",
        "Days_Remaining",
        "Renewal_Status",
        "Priority"
      ];
      rows =
        certificates
          .filter(c => {
            const d = days(c);
            return d !== null &&
              d <= 90;
          })
          .sort((a, b) =>
            (days(a) === null ? 99999 : days(a)) -
            (days(b) === null ? 99999 : days(b))
          )
          .map(c => [
            c.Certificate_ID || "",
            c.Employee_ID || "",
            employeeName(c),
            c.Certification_Name || "",
            c.Department || "",
            formatDate_(c.Expiry_Date),
            days(c),
            c.Renewal_Status || "",
            renewalPriority_(days(c))
          ]);
      break;

    case "EMPLOYEE_COMPLIANCE": {
      title =
        "Employee Compliance Report";
      headers = [
        "Employee_ID",
        "Employee_Name",
        "Department",
        "Total_Certifications",
        "Active",
        "Expired",
        "Expiring_Within_60_Days",
        "Compliance_Percent"
      ];

      rows =
        employees.map(e => {

          const own =
            certificates.filter(
              c =>
                String(c.Employee_ID) ===
                String(e.Employee_ID)
            );

          const activeCount =
            own.filter(
              c =>
                certStatus(c) ===
                "ACTIVE"
            ).length;

          const expiredCount =
            own.filter(
              c =>
                certStatus(c) ===
                "EXPIRED"
            ).length;

          const expiring60 =
            own.filter(c => {
              const d = days(c);
              return d !== null &&
                d >= 0 &&
                d <= 60;
            }).length;

          const compliance =
            own.length
              ? Math.round(
                  (
                    activeCount /
                    own.length
                  ) * 100
                )
              : 0;

          return [
            e.Employee_ID || "",
            e.Employee_Name || "",
            e.Department || "",
            own.length,
            activeCount,
            expiredCount,
            expiring60,
            compliance + "%"
          ];
        });
      break;
    }

    case "DEPARTMENT_COMPLIANCE": {
      title =
        "Department Compliance Report";
      headers = [
        "Department",
        "Total_Certifications",
        "Active",
        "Expired",
        "Expiring_Within_60_Days",
        "Compliance_Percent"
      ];

      const groups = {};

      certificates.forEach(c => {

        const department =
          c.Department ||
          (
            employeeMap[
              String(c.Employee_ID)
            ] || {}
          ).Department ||
          "Unknown";

        if (!groups[department]) {
          groups[department] = [];
        }

        groups[department].push(c);
      });

      rows =
        Object.keys(groups)
          .sort()
          .map(department => {

            const own =
              groups[department];

            const activeCount =
              own.filter(
                c =>
                  certStatus(c) ===
                  "ACTIVE"
              ).length;

            const expiredCount =
              own.filter(
                c =>
                  certStatus(c) ===
                  "EXPIRED"
              ).length;

            const expiring60 =
              own.filter(c => {
                const d = days(c);
                return d !== null &&
                  d >= 0 &&
                  d <= 60;
              }).length;

            const compliance =
              own.length
                ? Math.round(
                    (
                      activeCount /
                      own.length
                    ) * 100
                  )
                : 0;

            return [
              department,
              own.length,
              activeCount,
              expiredCount,
              expiring60,
              compliance + "%"
            ];
          });
      break;
    }

    case "MANDATORY_CERTIFICATION":
      title =
        "Mandatory Certification Report";
      headers = [
        "Certificate_ID",
        "Employee_ID",
        "Employee_Name",
        "Certification_Name",
        "Department",
        "Mandatory",
        "Expiry_Date",
        "Days_Remaining",
        "Status"
      ];
      rows =
        certificates
          .filter(c => {
            const v =
              String(
                c.Mandatory ||
                c.Is_Mandatory ||
                ""
              ).toUpperCase();

            return v === "TRUE" ||
              v === "YES";
          })
          .map(c => [
            c.Certificate_ID || "",
            c.Employee_ID || "",
            employeeName(c),
            c.Certification_Name || "",
            c.Department || "",
            "YES",
            formatDate_(c.Expiry_Date),
            days(c) === null
              ? ""
              : days(c),
            certStatus(c)
          ]);
      break;

    case "CERTIFICATION_TYPE":
      title =
        "Certification Type Report";
      headers = [
        "Certification_Category",
        "Count"
      ];
      {
        const map = {};

        certificates.forEach(c => {

          const key =
            c.Certification_Category ||
            "Unclassified";

          map[key] =
            (map[key] || 0) + 1;
        });

        rows =
          Object.keys(map)
            .sort()
            .map(k => [
              k,
              map[k]
            ]);
      }
      break;

    case "ISSUING_BODY":
      title =
        "Issuing Body Report";
      headers = [
        "Issuing_Body",
        "Count"
      ];
      {
        const map = {};

        certificates.forEach(c => {

          const key =
            c.Issuing_Body ||
            "Unknown";

          map[key] =
            (map[key] || 0) + 1;
        });

        rows =
          Object.keys(map)
            .sort()
            .map(k => [
              k,
              map[k]
            ]);
      }
      break;

    case "DOCUMENT_CONTROL":
      title =
        "Document Control Report";
      headers = [
        "Document_ID",
        "Certificate_ID",
        "Document_Name",
        "Classification",
        "Login_Required",
        "Status"
      ];
      rows =
        documents.map(d => [
          d.Document_ID || "",
          d.Certificate_ID || "",
          d.Document_Name || "",
          d.Classification || "",
          d.Login_Required || "",
          d.Status || ""
        ]);
      break;

    case "NOTIFICATION_HISTORY":
      title =
        "Notification History Report";
      headers = [
        "Notification_ID",
        "Certificate_ID",
        "Employee_ID",
        "Notification_Type",
        "Days_Before_Expiry",
        "To",
        "CC",
        "Sent_Date",
        "Status",
        "Error_Message"
      ];
      rows =
        notifications.map(n => [
          n.Notification_ID || "",
          n.Certificate_ID || "",
          n.Employee_ID || "",
          n.Notification_Type || "",
          n.Days_Before_Expiry || "",
          n.To ||
            n.Recipients ||
            "",
          n.CC ||
            n.Cc ||
            "",
          formatDateTime_(
            n.Sent_Date
          ),
          n.Status || "",
          n.Error_Message || ""
        ]);
      break;

    case "AUDIT_TRAIL":
      title =
        "Audit Trail Report";
      headers = [
        "Audit_ID",
        "Timestamp",
        "Email",
        "Role_ID",
        "Action",
        "Module",
        "Record_ID",
        "Description",
        "Result"
      ];
      rows =
        audit.map(a => [
          a.Audit_ID ||
            a.Log_ID ||
            "",
          formatDateTime_(
            a.Timestamp ||
            a.Date ||
            a.Created_Date
          ),
          a.Email ||
            a.User_Email ||
            "",
          a.Role_ID || "",
          a.Action || "",
          a.Module || "",
          a.Record_ID || "",
          a.Description || "",
          a.Result || ""
        ]);
      break;

    case "CERTIFICATION_GAP":
      title =
        "Certification Gap Report";
      headers = [
        "Employee_ID",
        "Employee_Name",
        "Department",
        "Certificate_ID",
        "Certification_Name",
        "Gap_Type",
        "Expiry_Date",
        "Days_Remaining"
      ];
      rows = [];

      certificates.forEach(c => {

        const d = days(c);

        if (
          certStatus(c) ===
          "EXPIRED"
        ) {
          rows.push([
            c.Employee_ID || "",
            employeeName(c),
            c.Department || "",
            c.Certificate_ID || "",
            c.Certification_Name || "",
            "EXPIRED_CERTIFICATION",
            formatDate_(
              c.Expiry_Date
            ),
            d
          ]);
        }
      });
      break;
    case "IMPORT_EXPORT_AUTHORIZATION": {
      title = "Import / Export Authorization Register";

      const authorizations =
        getSheetObjects_(
          CONFIG.SHEETS.IMPORT_EXPORT_AUTHORIZATIONS
        );

      const resolveAuthorizationOwner_ = function(a) {
        return (
          a.Responsible_Employee_Name ||
          (
            employeeMap[
              String(a.Responsible_Employee_ID || "")
            ] || {}
          ).Employee_Name ||
          a.Responsible_Employee_ID ||
          ""
        );
      };

      headers = [
        "Authorization_ID",
        "Authorization_Type",
        "Authorization_Name",
        "Authorization_Number",
        "Issuing_Authority",
        "Applicable_Activity",
        "Applicability",
        "Issue_Date",
        "Expiry_Date",
        "Responsible_Person",
        "Verification_Status",
        "Status",
        "Renewal_Status",
        "Document_Reference"
      ];

      rows = authorizations.map(function(a) {
        return [
          a.Authorization_ID || "",
          a.Authorization_Type || "",
          a.Authorization_Name || "",
          a.Authorization_Number || "",
          a.Issuing_Authority || "",
          a.Applicable_Activity || "",
          a.Applicability_Status || "",
          formatDate_(a.Issue_Date),
          formatDate_(a.Expiry_Date),
          resolveAuthorizationOwner_(a),
          a.Verification_Status || "",
          authorizationStatus_(a),
          authorizationRenewalStatus_(a),
          a.Document_ID || ""
        ];
      });

      /*
       * Additive E3.C1 report sections only.
       * Source: the existing 14_IMPORT_EXPORT_AUTHORIZATIONS register.
       * No new licence database or renewal engine is created.
       */
      const physicalRows = authorizations.map(function(a) {
        /*
         * The current E3.C1 authorization schema has no
         * Physical_Location field. Do not invent a location.
         * If a future/extended record contains an equivalent
         * location field, use it without changing the schema.
         */
        const physicalLocation =
          a.Physical_Location ||
          a.Document_Physical_Location ||
          a.Storage_Location ||
          "";

        return [
          a.Authorization_Name ||
            a.Authorization_Type ||
            "",
          a.Authorization_Number || "",
          a.Issuing_Authority || "",
          formatDate_(a.Issue_Date),
          formatDate_(a.Expiry_Date),
          resolveAuthorizationOwner_(a),
          physicalLocation,
          authorizationStatus_(a)
        ];
      });

      const digitalRows = authorizations.map(function(a) {
        const expiryDays =
          daysRemaining_(
            a.Expiry_Date
          );

        return [
          a.Authorization_Name ||
            a.Authorization_Type ||
            "",
          a.Authorization_Number || "",
          formatDate_(a.Expiry_Date),
          expiryDays === null
            ? ""
            : expiryDays,
          authorizationRenewalStatus_(a),
          resolveAuthorizationOwner_(a),
          a.Remarks || "-"
        ];
      });

      sections = [
        {
          title: "PHYSICAL REGISTER – LICENCE REGISTER",
          headers: [
            "Licence/Permit",
            "Reference No.",
            "Authority",
            "Issue Date",
            "Expiry Date",
            "Owner",
            "Physical Location",
            "Status"
          ],
          rows: physicalRows,
          emptyMessage:
            "No licence/authorization records available."
        },
        {
          title: "DIGITAL REGISTER – LICENCE EXPIRY TRACKER",
          headers: [
            "Licence",
            "Reference",
            "Expiry Date",
            "Days to Expiry",
            "Renewal Status",
            "Responsible Person",
            "Remarks"
          ],
          rows: digitalRows,
          emptyMessage:
            "No licence/authorization records available."
        }
      ];

      break;
    }


    case "MANAGEMENT_SUMMARY": {
      title =
        "Management Summary";
      headers = [
        "Metric",
        "Value"
      ];

      const activeCount =
        certificates.filter(
          c =>
            certStatus(c) ===
            "ACTIVE"
        ).length;

      const expiredCount =
        certificates.filter(
          c =>
            certStatus(c) ===
            "EXPIRED"
        ).length;

      const within60 =
        certificates.filter(c => {
          const d = days(c);
          return d !== null &&
            d >= 0 &&
            d <= 60;
        }).length;

      const within90 =
        certificates.filter(c => {
          const d = days(c);
          return d !== null &&
            d >= 0 &&
            d <= 90;
        }).length;

      const compliance =
        certificates.length
          ? Math.round(
              (
                activeCount /
                certificates.length
              ) * 100
            )
          : 0;

      rows = [
        [
          "Total Certifications",
          certificates.length
        ],
        [
          "Active Certifications",
          activeCount
        ],
        [
          "Expired Certifications",
          expiredCount
        ],
        [
          "Expiring Within 60 Days",
          within60
        ],
        [
          "Expiring Within 90 Days",
          within90
        ],
        [
          "Employees",
          employees.length
        ],
        [
          "Controlled Documents",
          documents.length
        ],
        [
          "Notifications Logged",
          notifications.length
        ],
        [
          "Audit Records",
          audit.length
        ],
        [
          "Compliance Percent",
          compliance + "%"
        ],
        [
          "Generated Date",
          Utilities.formatDate(
            new Date(),
            CONFIG.TIMEZONE,
            "dd-MMM-yyyy HH:mm:ss"
          )
        ]
      ];
      break;
    }

    default:
      throw new Error(
        "Unsupported report type: " +
        reportType
      );
  }

  return {
    title: title,
    headers: headers,
    rows: rows,
    sections: sections
  };
}


/************************************************************
 * HELPER: SHEET → OBJECTS
 ************************************************************/

function getSheetObjects_(
  sheetName
) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(
      "Sheet not found: " +
      sheetName
    );
  }

  const values =
    sheet.getDataRange().getValues();

  if (
    !values ||
    values.length < 2
  ) {
    return [];
  }

  const headers =
    values[0].map(
      h => String(h).trim()
    );

  return values
    .slice(1)
    .filter(
      row =>
        row.some(
          cell =>
            cell !== "" &&
            cell !== null
        )
    )
    .map(row => {

      const obj = {};

      headers.forEach(
        (header, index) => {

          obj[header] =
            row[index];

        }
      );

      return obj;
    });
}


/************************************************************
 * HELPER: FIND BY FIELD
 ************************************************************/

function findByField_(
  sheetName,
  field,
  value
) {

  const rows =
    getSheetObjects_(
      sheetName
    );

  return (
    rows.find(
      row =>
        String(row[field]) ===
        String(value)
    ) || null
  );
}


/************************************************************
 * USER LOOKUP
 ************************************************************/

function findUserByEmail_(
  email
) {

  if (!email) {
    return null;
  }

  const rows =
    getSheetObjects_(
      CONFIG.SHEETS.USERS
    );

  const normalized =
    String(email)
      .trim()
      .toLowerCase();

  return (
    rows.find(
      row =>
        String(row.Email)
          .trim()
          .toLowerCase() ===
        normalized
    ) || null
  );
}


/************************************************************
 * CURRENT USER EMAIL
 ************************************************************/

function getCurrentUserEmail_() {
  if (CCM_REQUEST_USER_EMAIL) return String(CCM_REQUEST_USER_EMAIL).trim().toLowerCase();
  try { return Session.getActiveUser().getEmail().trim().toLowerCase(); } catch (error) { return ""; }
}


/************************************************************
 * UPDATE LAST LOGIN
 ************************************************************/

function updateLastLogin_(
  userId
) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        CONFIG.SHEETS.USERS
      );

  const values =
    sheet.getDataRange().getValues();

  const headers =
    values[0];

  const idIndex =
    headers.indexOf("User_ID");

  const loginIndex =
    headers.indexOf("Last_Login");

  if (
    idIndex === -1 ||
    loginIndex === -1
  ) {
    return;
  }

  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    if (
      String(values[i][idIndex]) ===
      String(userId)
    ) {

      sheet
        .getRange(
          i + 1,
          loginIndex + 1
        )
        .setValue(
          new Date()
        );

      break;
    }
  }
}


/************************************************************
 * DATE CALCULATION
 ************************************************************/

function daysRemaining_(
  expiry
) {

  if (!expiry) {
    return null;
  }

  const expiryDate =
    startOfDay_(
      new Date(expiry)
    );

  if (
    isNaN(
      expiryDate.getTime()
    )
  ) {
    return null;
  }

  const today =
    startOfDay_(
      new Date()
    );

  return Math.ceil(
    (
      expiryDate.getTime() -
      today.getTime()
    ) /
    (
      1000 *
      60 *
      60 *
      24
    )
  );
}


function startOfDay_(
  date
) {

  const d =
    new Date(date);

  d.setHours(
    0,
    0,
    0,
    0
  );

  return d;
}


/************************************************************
 * RENEWAL STATUS
 ************************************************************/

function calculateRenewalStatus_(
  days
) {

  if (days === null) {
    return "UNKNOWN";
  }

  if (days < 0) {
    return "EXPIRED";
  }

  if (days <= 7) {
    return "CRITICAL";
  }

  if (days <= 15) {
    return "URGENT";
  }

  if (days <= 30) {
    return "RENEWAL_DUE";
  }

  if (days <= 60) {
    return "PLANNING";
  }

  if (days <= 90) {
    return "UPCOMING";
  }

  return "NOT_DUE";
}


function renewalPriority_(
  days
) {

  if (days < 0) {
    return "CRITICAL";
  }

  if (days <= 7) {
    return "CRITICAL";
  }

  if (days <= 15) {
    return "HIGH";
  }

  if (days <= 30) {
    return "MEDIUM";
  }

  return "LOW";
}


/************************************************************
 * DEPARTMENT COMPLIANCE
 ************************************************************/

function departmentCompliance_(
  certificates
) {

  const data = {};

  certificates.forEach(cert => {

    const department =
      cert.Department ||
      "Unassigned";

    if (!data[department]) {

      data[department] = {
        total: 0,
        compliant: 0
      };
    }

    data[department].total++;

    const days =
      daysRemaining_(
        cert.Expiry_Date
      );

    if (
      days !== null &&
      days > 30
    ) {

      data[department].compliant++;
    }
  });

  Object.keys(data)
    .forEach(department => {

      const d =
        data[department];

      d.compliance =
        d.total
          ? Math.round(
              (
                d.compliant /
                d.total
              ) * 10000
            ) / 100
          : 0;
    });

  return data;
}


/************************************************************
 * EMPLOYEE CERTIFICATION SUMMARY
 ************************************************************/

function employeeCertificationSummary_(
  certificates
) {

  let active = 0;
  let expiring = 0;
  let expired = 0;

  certificates.forEach(cert => {

    const days =
      daysRemaining_(
        cert.Expiry_Date
      );

    if (days < 0) {
      expired++;
    } else {
      active++;
    }

    if (
      days !== null &&
      days <= 90
    ) {
      expiring++;
    }
  });

  return {
    total: certificates.length,
    active: active,
    expiring: expiring,
    expired: expired
  };
}


/************************************************************
 * RECENT RENEWALS
 ************************************************************/

function getRecentRenewals_() {

  const certificates =
    getSheetObjects_(
      CONFIG.SHEETS.CERTIFICATIONS
    );

  return certificates
    .filter(
      cert =>
        String(cert.Renewal_Status)
          .toUpperCase() ===
        "RENEWED"
    )
    .slice(-10)
    .reverse()
    .map(
      sanitizeCertificate_
    );
}


/************************************************************
 * NOTIFICATION HISTORY
 ************************************************************/

function getNotificationHistory_(
  certificateId
) {

  return getSheetObjects_(
    CONFIG.SHEETS.NOTIFICATION_LOG
  )
  .filter(
    x =>
      String(x.Certificate_ID) ===
      String(certificateId)
  )
  .map(x => ({
    Notification_ID:
      x.Notification_ID,

    Notification_Type:
      x.Notification_Type,

    Days_Before_Expiry:
      x.Days_Before_Expiry,

    Sent_Date:
      formatDateTime_(x.Sent_Date),

    Status:
      x.Status
  }));
}


/************************************************************
 * AUDIT HISTORY
 ************************************************************/

function getAuditHistory_(
  recordId
) {

  return getSheetObjects_(
    CONFIG.SHEETS.AUDIT_LOG
  )
  .filter(
    x =>
      String(x.Record_ID) ===
      String(recordId)
  )
  .slice(-50)
  .reverse();
}


/************************************************************
 * CERTIFICATE → EMPLOYEE
 ************************************************************/

function getCertificateEmployeeId_(
  certificateId
) {

  const certificate =
    findByField_(
      CONFIG.SHEETS.CERTIFICATIONS,
      "Certificate_ID",
      certificateId
    );

  return certificate
    ? certificate.Employee_ID
    : "";
}


/************************************************************
 * SANITIZERS
 ************************************************************/

function sanitizeUser_(
  user
) {

  return {
    User_ID: user.User_ID,
    Email: user.Email,
    Name: user.Name,
    Employee_ID: user.Employee_ID,
    Department: user.Department,
    Role_ID: user.Role_ID,
    Status: user.Status
  };
}


function sanitizeCertificate_(
  certificate
) {

  return {

    Certificate_ID:
      certificate.Certificate_ID,

    Employee_ID:
      certificate.Employee_ID,

    Employee_Name:
      certificate.Employee_Name,

    Department:
      certificate.Department,

    Designation:
      certificate.Designation,

    Certification_Name:
      certificate.Certification_Name,

    Certification_Category:
      certificate.Certification_Category,

    Issuing_Body:
      certificate.Issuing_Body,

    Certificate_Number:
      certificate.Certificate_Number,

    Issue_Date:
      formatDate_(
        certificate.Issue_Date
      ),

    Expiry_Date:
      formatDate_(
        certificate.Expiry_Date
      ),

    Mandatory:
      certificate.Mandatory,

    Status:
      certificate.Status,

    Renewal_Status:
      certificate.Renewal_Status,

    Document_ID:
      certificate.Document_ID,

    Remarks:
      certificate.Remarks
  };
}


function sanitizeEmployee_(
  employee
) {

  return {

    Employee_ID:
      employee.Employee_ID,

    Employee_Name:
      employee.Employee_Name,

    Email:
      employee.Email,

    Department:
      employee.Department,

    Designation:
      employee.Designation,

    Location:
      employee.Location,

    Employment_Status:
      employee.Employment_Status,

    Joining_Date:
      formatDate_(
        employee.Joining_Date
      )
  };
}


/************************************************************
 * FORMATTERS
 ************************************************************/

function formatDate_(
  value
) {

  if (!value) {
    return "";
  }

  const date =
    new Date(value);

  if (
    isNaN(
      date.getTime()
    )
  ) {

    return String(value);
  }

  return Utilities.formatDate(
    date,
    CONFIG.TIMEZONE,
    "dd-MMM-yyyy"
  );
}


function formatDateTime_(
  value
) {

  if (!value) {
    return "";
  }

  const date =
    new Date(value);

  if (
    isNaN(
      date.getTime()
    )
  ) {

    return String(value);
  }

  return Utilities.formatDate(
    date,
    CONFIG.TIMEZONE,
    "dd-MMM-yyyy HH:mm:ss"
  );
}


/************************************************************
 * ID GENERATOR
 ************************************************************/

function generateId_(
  prefix
) {

  return (
    prefix +
    "-" +
    Utilities.formatDate(
      new Date(),
      CONFIG.TIMEZONE,
      "yyyyMMddHHmmss"
    ) +
    "-" +
    Math.floor(
      Math.random() * 1000
    )
  );
}


/************************************************************
 * AUDIT LOG
 ************************************************************/

function logAudit_(
  email,
  action,
  module,
  recordId,
  description,
  result
) {

  try {

    const sheet =
      SpreadsheetApp
        .getActiveSpreadsheet()
        .getSheetByName(
          CONFIG.SHEETS.AUDIT_LOG
        );

    if (!sheet) {
      return;
    }

    const user =
      findUserByEmail_(
        email
      );

    const row = [

      generateId_("AUD"),

      new Date(),

      email || "",

      user
        ? user.Role_ID
        : "",

      action,

      module,

      recordId || "",

      description || "",

      "",

      result

    ];

    sheet
      .getRange(
        sheet.getLastRow() + 1,
        1,
        1,
        row.length
      )
      .setValues([row]);

  } catch (error) {

    console.error(
      "Audit logging failed: " +
      error.message
    );
  }
}


/************************************************************
 * JSON RESPONSE
 ************************************************************/

function jsonResponse_(
  object
) {

  return ContentService
    .createTextOutput(
      JSON.stringify(
        object,
        null,
        2
      )
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );
}


/************************************************************
 * MANUAL TEST FUNCTION
 ************************************************************/

function testBackend() {

  const result =
    apiHealth_();

  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );
}


/************************************************************
 * CREATE DAILY RENEWAL TRIGGER
 ************************************************************/

function createDailyTrigger() {

  const triggers =
    ScriptApp.getProjectTriggers();

  triggers.forEach(
    trigger => {

      if (
        trigger.getHandlerFunction() ===
        "runRenewalCheck"
      ) {

        ScriptApp.deleteTrigger(
          trigger
        );
      }
    }
  );

  ScriptApp
    .newTrigger(
      "runRenewalCheck"
    )
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  Logger.log(
    "Daily renewal trigger created."
  );
}

/************************************************************
 * CCM 12.0 — RENEWAL RULE SETUP
 *
 * Ensures the mandatory 60-day formal renewal rule exists.
 * Run once after installing CCM 12.0.
 ************************************************************/

function setupCCM12RenewalRules() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        CONFIG.SHEETS.RENEWAL_RULES
      );

  if (!sheet) {
    throw new Error(
      "Sheet not found: " +
      CONFIG.SHEETS.RENEWAL_RULES
    );
  }

  const values =
    sheet.getDataRange().getValues();

  if (!values.length) {
    throw new Error(
      "05_RENEWAL_RULES has no header row."
    );
  }

  const headers =
    values[0].map(h => String(h).trim());

  const required = [
    "Rule_ID",
    "Days_Before_Expiry",
    "Email_Type",
    "Recipient",
    "CC_Recipient",
    "Enabled"
  ];

  const missing =
    required.filter(h => headers.indexOf(h) === -1);

  if (missing.length) {
    throw new Error(
      "05_RENEWAL_RULES missing columns: " +
      missing.join(", ")
    );
  }

  const rows = values.slice(1);
  const daysCol = headers.indexOf("Days_Before_Expiry");
  const enabledCol = headers.indexOf("Enabled");

  const rules = [
    { id: "RULE-90", days: 90, type: "RENEWAL_90_DAYS", recipient: "EMPLOYEE", cc: "HR;MANAGER", priority: "MEDIUM", description: "Renewal planning notification 90 days before expiry" },
    { id: "RULE-60", days: 60, type: "RENEWAL_60_DAYS", recipient: "EMPLOYEE", cc: "HR;MANAGER", priority: "HIGH", description: "Formal renewal notification 60 days before expiry" },
    { id: "RULE-30", days: 30, type: "RENEWAL_30_DAYS", recipient: "EMPLOYEE", cc: "HR;MANAGER", priority: "HIGH", description: "Renewal reminder 30 days before expiry" },
    { id: "RULE-15", days: 15, type: "RENEWAL_15_DAYS", recipient: "EMPLOYEE", cc: "HR;MANAGER", priority: "HIGH", description: "Renewal escalation 15 days before expiry" },
    { id: "RULE-7", days: 7, type: "RENEWAL_7_DAYS", recipient: "EMPLOYEE", cc: "HR;MANAGER", priority: "URGENT", description: "Urgent renewal escalation 7 days before expiry" },
    { id: "RULE-0", days: 0, type: "RENEWAL_0_DAYS", recipient: "EMPLOYEE", cc: "HR;MANAGER", priority: "URGENT", description: "Certification expires today notification" },
    { id: "RULE-EXPIRED", days: -1, type: "CERTIFICATION_EXPIRED", recipient: "EMPLOYEE", cc: "HR;MANAGER", priority: "CRITICAL", description: "Certification expired escalation" }
  ];

  const created = [];

  rules.forEach(def => {
    const exists = rows.some(row =>
      Number(row[daysCol]) === def.days &&
      String(row[enabledCol]).toUpperCase() === "TRUE"
    );

    if (exists) return;

    const row = headers.map(header => {
      switch (header) {
        case "Rule_ID": return def.id;
        case "Days_Before_Expiry": return def.days;
        case "Email_Type": return def.type;
        case "Recipient": return def.recipient;
        case "CC_Recipient": return def.cc;
        case "Enabled": return "TRUE";
        case "Priority": return def.priority;
        case "Description": return def.description;
        default: return "";
      }
    });

    sheet
      .getRange(sheet.getLastRow() + 1, 1, 1, row.length)
      .setValues([row]);

    created.push(def.id);

    logAudit_(
      getCurrentUserEmail_(),
      "RENEWAL_RULE_CREATED",
      "RENEWALS",
      def.id,
      def.description,
      "SUCCESS"
    );
  });

  return {
    success: true,
    created: created,
    message: created.length
      ? "CCM 12.0 renewal threshold rules configured."
      : "All CCM 12.0 renewal threshold rules already exist."
  };
}


/************************************************************
 * CCM 12.0 — RENEWAL EVALUATION TEST
 *
 * Does NOT send mail.
 ************************************************************/

function testCCM12RenewalEvaluation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const certificates = getSheetObjects_(CONFIG.SHEETS.CERTIFICATIONS);
  const rules = getSheetObjects_(CONFIG.SHEETS.RENEWAL_RULES);

  const evaluations = [];
  let triggered=0, noRule=0, invalidExpiry=0;

  certificates.forEach(function(cert) {
    const days = daysRemaining_(cert.Expiry_Date);

    if (days === null) {
      invalidExpiry++;
      evaluations.push({
        Certificate_ID:cert.Certificate_ID||"",
        Employee_ID:cert.Employee_ID||"",
        Certification_Name:cert.Certification_Name||"",
        Expiry_Date:cert.Expiry_Date||"",
        Days_Remaining:"",
        Rule_ID:"",
        Email_Type:"",
        Recipient:"",
        CC_Recipient:"",
        Trigger:false,
        Result:"INVALID_EXPIRY"
      });
      return;
    }

    const rule=findRenewalRule_(rules,days);
    if (rule) triggered++; else noRule++;

    evaluations.push({
      Certificate_ID:cert.Certificate_ID||"",
      Employee_ID:cert.Employee_ID||"",
      Certification_Name:cert.Certification_Name||"",
      Expiry_Date:formatDate_(cert.Expiry_Date),
      Days_Remaining:days,
      Rule_ID:rule ? rule.Rule_ID||"" : "",
      Email_Type:rule ? rule.Email_Type||"" : "",
      Recipient:rule ? rule.Recipient||"" : "",
      CC_Recipient:rule ? rule.CC_Recipient||"" : "",
      Trigger:!!rule,
      Result:rule ? "TRIGGERED" : "NO_RULE"
    });
  });

  const sheetName="CCM 12.0 — RENEWAL EVALUATION";
  let sheet=ss.getSheetByName(sheetName);
  if (!sheet) sheet=ss.insertSheet(sheetName);
  sheet.clearContents();
  sheet.clearFormats();

  const summary=[
    ["CCM 12.0 — RENEWAL EVALUATION",""],
    ["Generated At",new Date()],
    ["Checked",certificates.length],
    ["Triggered",triggered],
    ["No Rule",noRule],
    ["Invalid Expiry",invalidExpiry],
    ["Mail Sent","NO — TEST ONLY"],
    ["",""]
  ];

  const headers=[
    "Certificate ID","Employee ID","Certification","Expiry Date",
    "Days Remaining","Rule ID","Email Type","Recipient","CC",
    "Trigger","Result"
  ];

  const rows=evaluations.map(function(x){
    return [
      x.Certificate_ID,x.Employee_ID,x.Certification_Name,x.Expiry_Date,
      x.Days_Remaining,x.Rule_ID,x.Email_Type,x.Recipient,
      x.CC_Recipient,x.Trigger?"YES":"NO",x.Result
    ];
  });

  sheet.getRange(1,1,summary.length,2).setValues(summary);
  sheet.getRange(summary.length+1,1,1,headers.length).setValues([headers]);
  if(rows.length) sheet.getRange(summary.length+2,1,rows.length,headers.length).setValues(rows);

  sheet.getRange(1,1,1,2).setFontWeight("bold");
  sheet.getRange(summary.length+1,1,1,headers.length).setFontWeight("bold");
  sheet.getDataRange().setWrap(true);
  sheet.setFrozenRows(summary.length+1);
  sheet.autoResizeColumns(1,headers.length);

  const result={
    success:true,checked:certificates.length,triggered:triggered,
    noRule:noRule,invalidExpiry:invalidExpiry,evaluations:evaluations,
    mailSent:false,resultSheet:sheetName
  };
  Logger.log(JSON.stringify(result,null,2));
  console.log(JSON.stringify(result,null,2));
  return result;
}


/************************************************************
 * MANUAL AUDIT TEST
 ************************************************************/
function testLogAudit() {

  logAudit_(
    getCurrentUserEmail_(),
    "TEST",
    "SYSTEM",
    "TEST-001",
    "Audit logging test",
    "SUCCESS"
  );

  return {
    success: true,
    message: "Audit log test completed."
  };
}


/**
 * CCM 11I — PRODUCTION API GATEWAY
 *
 * Replace ONLY the existing ccmExecute() function with this version.
 * Keep the rest of the backend unchanged.
 *
 * Exposes the existing server-side functions through Apps Script
 * Execution API (scripts.run), including the functional write actions.
 */

/************************************************************
 * E3.C1 — IMPORT / EXPORT LICENSE & AUTHORIZATION CONTROL
 *
 * This extends the existing CCM architecture. It does not
 * replace the certificate renewal engine, document control,
 * RBAC, notification log or audit mechanism.
 ************************************************************/

const E3C1_AUTHORIZATION_HEADERS = [
  "Authorization_ID",
  "Authorization_Type",
  "Authorization_Name",
  "Authorization_Number",
  "Issuing_Authority",
  "Applicable_Activity",
  "Applicable_Country",
  "Applicable_Process",
  "Applicability_Status",
  "Issue_Date",
  "Expiry_Date",
  "Renewal_Lead_Days",
  "Responsible_Department",
  "Responsible_Employee_ID",
  "Responsible_Employee_Name",
  "Verification_Status",
  "Verification_Date",
  "Verified_By",
  "Status",
  "Renewal_Status",
  "Document_ID",
  "Remarks",
  "Created_By",
  "Created_Date",
  "Last_Modified_By",
  "Last_Modified_Date",
  "Procedure_Document_ID"
];

const E3C1_AUTHORIZATION_TYPES = [
  "Import License",
  "Export License",
  "Import Authorization",
  "Export Authorization",
  "Customs Authorization",
  "Regulatory Authorization",
  "Government Registration",
  "Permit",
  "Certificate / Authorization",
  "Other"
];

function setupImportExportAuthorizationControl_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEETS.IMPORT_EXPORT_AUTHORIZATIONS);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.IMPORT_EXPORT_AUTHORIZATIONS);
  }

  const existing = sheet.getLastColumn() > 0
    ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String)
    : [];

  if (existing.length === 0 || existing.every(function(v) { return !String(v).trim(); })) {
    sheet.getRange(1, 1, 1, E3C1_AUTHORIZATION_HEADERS.length)
      .setValues([E3C1_AUTHORIZATION_HEADERS]);
  } else {
    E3C1_AUTHORIZATION_HEADERS.forEach(function(header) {
      if (existing.indexOf(header) === -1) {
        sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
      }
    });
  }

  sheet.setFrozenRows(1);

  const permissionRows = [
    ["PERM-AUTH-VIEW", "View Import / Export Authorizations"],
    ["PERM-AUTH-CREATE", "Create Import / Export Authorization"],
    ["PERM-AUTH-EDIT", "Edit Import / Export Authorization"],
    ["PERM-AUTH-VERIFY", "Verify / Reject Import / Export Authorization"],
    ["PERM-AUTH-RENEW", "Manage Authorization Renewal"],
    ["PERM-AUTH-DOCUMENT", "Access Authorization Documents"],
    ["PERM-AUTH-REPORT", "Generate Authorization Reports"]
  ];

  const permissionSheet = ss.getSheetByName(CONFIG.SHEETS.PERMISSIONS);
  if (permissionSheet) {
    const values = permissionSheet.getDataRange().getValues();
    const headers = values.length ? values[0].map(String) : [];
    const idIndex = headers.indexOf("Permission_ID");
    const nameIndex = headers.indexOf("Permission_Name");
    if (idIndex !== -1) {
      permissionRows.forEach(function(row) {
        const exists = values.slice(1).some(function(r) {
          return String(r[idIndex]).trim() === row[0];
        });
        if (!exists) {
          const newRow = new Array(Math.max(headers.length, 2)).fill("");
          newRow[idIndex] = row[0];
          if (nameIndex !== -1) newRow[nameIndex] = row[1];
          permissionSheet.appendRow(newRow);
        }
      });
    }
  }

  const rolePermissionSheet = ss.getSheetByName(CONFIG.SHEETS.ROLE_PERMISSIONS);
  const rolesSheet = ss.getSheetByName(CONFIG.SHEETS.ROLES);
  if (rolePermissionSheet && rolesSheet) {
    const roleValues = rolesSheet.getDataRange().getValues();
    const roleHeaders = roleValues.length ? roleValues[0].map(String) : [];
    const roleIdIndex = roleHeaders.indexOf("Role_ID");
    const roleNameIndex = roleHeaders.indexOf("Role_Name");
    let adminRoleId = "";
    if (roleIdIndex !== -1) {
      for (let r = 1; r < roleValues.length; r++) {
        const roleName = roleNameIndex !== -1 ? String(roleValues[r][roleNameIndex] || "").toUpperCase() : "";
        const roleId = String(roleValues[r][roleIdIndex] || "");
        if (roleName.indexOf("ADMIN") !== -1 || roleId.toUpperCase().indexOf("ADMIN") !== -1) {
          adminRoleId = roleId;
          break;
        }
      }
    }

    if (adminRoleId) {
      const rpValues = rolePermissionSheet.getDataRange().getValues();
      const rpHeaders = rpValues.length ? rpValues[0].map(String) : [];
      const rpRoleIndex = rpHeaders.indexOf("Role_ID");
      const rpPermissionIndex = rpHeaders.indexOf("Permission_ID");
      const rpAllowedIndex = rpHeaders.indexOf("Allowed");

      if (rpRoleIndex !== -1 && rpPermissionIndex !== -1) {
        permissionRows.forEach(function(p) {
          const exists = rpValues.slice(1).some(function(row) {
            return String(row[rpRoleIndex] || "") === adminRoleId &&
                   String(row[rpPermissionIndex] || "") === p[0];
          });
          if (!exists) {
            const newRow = new Array(rpHeaders.length).fill("");
            newRow[rpRoleIndex] = adminRoleId;
            newRow[rpPermissionIndex] = p[0];
            if (rpAllowedIndex !== -1) newRow[rpAllowedIndex] = true;
            rolePermissionSheet.appendRow(newRow);
          }
        });
      }
    }
  }

  logAudit_(
    getCurrentUserEmail_(),
    "SETUP",
    "AUTHORIZATION_CONTROL",
    CONFIG.SHEETS.IMPORT_EXPORT_AUTHORIZATIONS,
    "E3.C1 Import / Export Authorization control initialized",
    "SUCCESS"
  );

  return {
    success: true,
    sheet: CONFIG.SHEETS.IMPORT_EXPORT_AUTHORIZATIONS,
    headers: E3C1_AUTHORIZATION_HEADERS,
    authorizationTypes: E3C1_AUTHORIZATION_TYPES,
    permissions: permissionRows.map(function(r) { return r[0]; })
  };
}

function authorizationStatus_(authorization) {
  const expiry = parseDate_(authorization.Expiry_Date);
  const source = String(authorization.Status || "").trim().toUpperCase();

  if (source === "SUSPENDED" || source === "CANCELLED" || source === "INACTIVE") {
    return source;
  }

  if (!expiry) return source || "INACTIVE";

  const days = daysRemaining_(authorization.Expiry_Date);
  if (days < 0) return "EXPIRED";
  if (days <= 90) return "EXPIRING";
  return "ACTIVE";
}

function authorizationRenewalStatus_(authorization) {
  const expiry = parseDate_(authorization.Expiry_Date);
  if (!expiry) return "NOT_CONFIGURED";

  const days = daysRemaining_(authorization.Expiry_Date);
  const lead = Number(authorization.Renewal_Lead_Days);

  if (days < 0) return "EXPIRED";
  if (isFinite(lead) && days <= lead) return "RENEWAL_DUE";
  if (days <= 90) return "UPCOMING";
  return "NOT_DUE";
}

function validateAuthorizationData_(data, existing) {
  data = data || {};
  const merged = Object.assign({}, existing || {}, data);

  const required = [
    ["Authorization_Type", "Authorization Type"],
    ["Authorization_Name", "Authorization Name"],
    ["Issuing_Authority", "Issuing Authority"],
    ["Applicable_Activity", "Applicable Activity"],
    ["Applicability_Status", "Applicability Status"],
    ["Issue_Date", "Issue Date"],
    ["Expiry_Date", "Expiry Date"],
    ["Responsible_Department", "Responsible Department"],
    ["Responsible_Employee_ID", "Responsible Employee"]
  ];

  for (let i = 0; i < required.length; i++) {
    const key = required[i][0];
    if (!String(merged[key] || "").trim()) {
      throw new Error(required[i][1] + " is required.");
    }
  }

  const applicability = String(merged.Applicability_Status).trim();
  if (["Applicable", "Not Applicable", "Under Review"].indexOf(applicability) === -1) {
    throw new Error("Applicability Status must be Applicable, Not Applicable or Under Review.");
  }

  if (applicability === "Not Applicable" && !String(merged.Remarks || "").trim()) {
    throw new Error("Remarks / justification is required when Applicability Status is Not Applicable.");
  }

  const issue = parseDate_(merged.Issue_Date);
  const expiry = parseDate_(merged.Expiry_Date);

  if (!issue || !expiry) {
    throw new Error("Issue Date and Expiry Date must be valid dates.");
  }

  if (expiry < issue) {
    throw new Error("Expiry Date cannot be earlier than Issue Date.");
  }

  const verification = String(
    merged.Verification_Status || "Pending"
  ).trim();

  if (["Pending", "Verified", "Rejected"].indexOf(verification) === -1) {
    throw new Error("Verification Status must be Pending, Verified or Rejected.");
  }

  return merged;
}

function apiAuthorizations_(status, applicability) {
  const auth = requirePermission_("PERM-AUTH-VIEW");
  if (!auth.allowed) return auth.response;

  let records = getSheetObjects_(CONFIG.SHEETS.IMPORT_EXPORT_AUTHORIZATIONS);

  records = records.map(function(row) {
    row.Status = authorizationStatus_(row);
    row.Renewal_Status = authorizationRenewalStatus_(row);
    return sanitizeAuthorization_(row);
  });

  if (status) {
    records = records.filter(function(row) {
      return String(row.Status).toUpperCase() === String(status).toUpperCase();
    });
  }

  if (applicability) {
    records = records.filter(function(row) {
      return String(row.Applicability_Status).toLowerCase() === String(applicability).toLowerCase();
    });
  }

  return {
    success: true,
    count: records.length,
    data: records,
    types: E3C1_AUTHORIZATION_TYPES
  };
}

function apiAuthorization360_(authorizationId) {
  const auth = requirePermission_("PERM-AUTH-VIEW");
  if (!auth.allowed) return auth.response;

  const row = findByField_(
    CONFIG.SHEETS.IMPORT_EXPORT_AUTHORIZATIONS,
    "Authorization_ID",
    authorizationId
  );

  if (!row) {
    return {
      success: false,
      error: "AUTHORIZATION_NOT_FOUND"
    };
  }

  row.Status = authorizationStatus_(row);
  row.Renewal_Status = authorizationRenewalStatus_(row);

  const documents = getSheetObjects_(CONFIG.SHEETS.DOCUMENTS)
    .filter(function(d) {
      return String(d.Document_ID || "") === String(row.Document_ID || "") ||
             String(d.Authorization_ID || "") === String(authorizationId);
    })
    .map(sanitizeDocumentForUser_);

  return {
    success: true,
    authorization: sanitizeAuthorization_(row),
    documents: documents,
    audit: getAuditHistory_(authorizationId)
  };
}

function createAuthorization_(data) {
  const auth = requirePermission_("PERM-AUTH-CREATE");
  if (!auth.allowed) return auth.response;

  const validated = validateAuthorizationData_(data || {});
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEETS.IMPORT_EXPORT_AUTHORIZATIONS);

  if (!sheet) {
    throw new Error("Sheet not found: " + CONFIG.SHEETS.IMPORT_EXPORT_AUTHORIZATIONS);
  }

  const now = new Date();
  const email = getCurrentUserEmail_();
  const id = generateId_("AUTH");

  const row = [
    id,
    validated.Authorization_Type || "",
    validated.Authorization_Name || "",
    validated.Authorization_Number || "",
    validated.Issuing_Authority || "",
    validated.Applicable_Activity || "",
    validated.Applicable_Country || "",
    validated.Applicable_Process || "",
    validated.Applicability_Status || "",
    validated.Issue_Date || "",
    validated.Expiry_Date || "",
    validated.Renewal_Lead_Days || 90,
    validated.Responsible_Department || "",
    validated.Responsible_Employee_ID || "",
    validated.Responsible_Employee_Name || "",
    validated.Verification_Status || "Pending",
    validated.Verification_Date || "",
    validated.Verified_By || "",
    authorizationStatus_(validated),
    authorizationRenewalStatus_(validated),
    validated.Document_ID || "",
    validated.Remarks || "",
    email,
    now,
    email,
    now,
    validated.Procedure_Document_ID || ""
  ];

  sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);

  logAudit_(
    email,
    "CREATE",
    "AUTHORIZATION",
    id,
    "Import / Export Authorization created",
    "SUCCESS"
  );

  return {
    success: true,
    authorizationId: id
  };
}

function updateAuthorization_(authorizationId, data) {
  const auth = requirePermission_("PERM-AUTH-EDIT");
  if (!auth.allowed) return auth.response;

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEETS.IMPORT_EXPORT_AUTHORIZATIONS);

  if (!sheet) throw new Error("Sheet not found: " + CONFIG.SHEETS.IMPORT_EXPORT_AUTHORIZATIONS);

  const values = sheet.getDataRange().getValues();
  if (!values.length) throw new Error("Authorization sheet has no headers.");

  const headers = values[0].map(String);
  const idIndex = headers.indexOf("Authorization_ID");
  const rowIndex = values.findIndex(function(row, index) {
    return index > 0 && String(row[idIndex]) === String(authorizationId);
  });

  if (rowIndex === -1) {
    return {
      success: false,
      error: "AUTHORIZATION_NOT_FOUND"
    };
  }

  const existing = {};
  headers.forEach(function(h, i) {
    existing[h] = values[rowIndex][i];
  });

  const merged = validateAuthorizationData_(data || {}, existing);

  Object.keys(data || {}).forEach(function(key) {
    const i = headers.indexOf(key);
    if (i !== -1) values[rowIndex][i] = data[key];
  });

  const statusIndex = headers.indexOf("Status");
  const renewalIndex = headers.indexOf("Renewal_Status");
  const modifiedByIndex = headers.indexOf("Last_Modified_By");
  const modifiedDateIndex = headers.indexOf("Last_Modified_Date");

  if (statusIndex !== -1) values[rowIndex][statusIndex] = authorizationStatus_(merged);
  if (renewalIndex !== -1) values[rowIndex][renewalIndex] = authorizationRenewalStatus_(merged);
  if (modifiedByIndex !== -1) values[rowIndex][modifiedByIndex] = getCurrentUserEmail_();
  if (modifiedDateIndex !== -1) values[rowIndex][modifiedDateIndex] = new Date();

  sheet.getRange(rowIndex + 1, 1, 1, values[rowIndex].length)
    .setValues([values[rowIndex]]);

  logAudit_(
    getCurrentUserEmail_(),
    "UPDATE",
    "AUTHORIZATION",
    authorizationId,
    "Import / Export Authorization updated",
    "SUCCESS"
  );

  return {
    success: true,
    authorizationId: authorizationId
  };
}

function verifyAuthorization_(authorizationId, verificationStatus, remarks) {
  const auth = requirePermission_("PERM-AUTH-VERIFY");
  if (!auth.allowed) return auth.response;

  const status = String(verificationStatus || "").trim();
  if (["Verified", "Rejected", "Pending"].indexOf(status) === -1) {
    return {
      success: false,
      error: "INVALID_VERIFICATION_STATUS"
    };
  }

  const row = findByField_(
    CONFIG.SHEETS.IMPORT_EXPORT_AUTHORIZATIONS,
    "Authorization_ID",
    authorizationId
  );

  if (!row) {
    return {
      success: false,
      error: "AUTHORIZATION_NOT_FOUND"
    };
  }

  const data = {
    Verification_Status: status,
    Verification_Date: new Date(),
    Verified_By: getCurrentUserEmail_()
  };

  if (remarks !== undefined && remarks !== null && String(remarks).trim()) {
    data.Remarks = remarks;
  }

  const result = updateAuthorization_(authorizationId, data);

  logAudit_(
    getCurrentUserEmail_(),
    status === "Rejected" ? "REJECT" : "VERIFY",
    "AUTHORIZATION",
    authorizationId,
    "Authorization verification status changed to " + status,
    "SUCCESS"
  );

  return result;
}

function apiAuthorizationRenewals_() {
  const auth = requirePermission_("PERM-AUTH-RENEW");
  if (!auth.allowed) return auth.response;

  const records = getSheetObjects_(
    CONFIG.SHEETS.IMPORT_EXPORT_AUTHORIZATIONS
  ).map(function(row) {
    row.Status = authorizationStatus_(row);
    row.Renewal_Status = authorizationRenewalStatus_(row);
    return sanitizeAuthorization_(row);
  }).filter(function(row) {
    return ["RENEWAL_DUE", "EXPIRED", "UPCOMING"].indexOf(
      String(row.Renewal_Status).toUpperCase()
    ) !== -1;
  });

  return {
    success: true,
    count: records.length,
    data: records
  };
}

function sanitizeAuthorization_(row) {
  return {
    Authorization_ID: row.Authorization_ID || "",
    Authorization_Type: row.Authorization_Type || "",
    Authorization_Name: row.Authorization_Name || "",
    Authorization_Number: row.Authorization_Number || "",
    Issuing_Authority: row.Issuing_Authority || "",
    Applicable_Activity: row.Applicable_Activity || "",
    Applicable_Country: row.Applicable_Country || "",
    Applicable_Process: row.Applicable_Process || "",
    Applicability_Status: row.Applicability_Status || "",
    Issue_Date: formatDate_(row.Issue_Date),
    Expiry_Date: formatDate_(row.Expiry_Date),
    Renewal_Lead_Days: row.Renewal_Lead_Days || "",
    Responsible_Department: row.Responsible_Department || "",
    Responsible_Employee_ID: row.Responsible_Employee_ID || "",
    Responsible_Employee_Name: row.Responsible_Employee_Name || "",
    Verification_Status: row.Verification_Status || "Pending",
    Verification_Date: formatDate_(row.Verification_Date),
    Verified_By: row.Verified_By || "",
    Status: authorizationStatus_(row),
    Renewal_Status: authorizationRenewalStatus_(row),
    Document_ID: row.Document_ID || "",
    Remarks: row.Remarks || "",
    Procedure_Document_ID: row.Procedure_Document_ID || "",
    Created_By: row.Created_By || "",
    Created_Date: formatDateTime_(row.Created_Date),
    Last_Modified_By: row.Last_Modified_By || "",
    Last_Modified_Date: formatDateTime_(row.Last_Modified_Date)
  };
}

function getAuthorizationDashboardKpis_() {
  const records = getSheetObjects_(
    CONFIG.SHEETS.IMPORT_EXPORT_AUTHORIZATIONS
  );

  let active = 0;
  let expiring = 0;
  let expired = 0;
  let pendingVerification = 0;
  let renewalDue = 0;
  let notApplicable = 0;

  records.forEach(function(row) {
    const status = authorizationStatus_(row);
    const renewal = authorizationRenewalStatus_(row);
    const applicability = String(row.Applicability_Status || "").trim();

    if (status === "ACTIVE") active++;
    if (status === "EXPIRING") expiring++;
    if (status === "EXPIRED") expired++;
    if (String(row.Verification_Status || "Pending") === "Pending") pendingVerification++;
    if (renewal === "RENEWAL_DUE") renewalDue++;
    if (applicability === "Not Applicable") notApplicable++;
  });

  return {
    total: records.length,
    active: active,
    expiring: expiring,
    expired: expired,
    pendingVerification: pendingVerification,
    renewalDue: renewalDue,
    notApplicable: notApplicable
  };
}

function getE3C1ProcedureFramework_() {
  return [
    "Identify applicable license/authorization requirements.",
    "Determine applicability to the company's import/export activity.",
    "Obtain the required license/authorization from the applicable authority.",
    "Verify validity and authenticity.",
    "Register the authorization in CCM.",
    "Upload and link supporting documentation.",
    "Assign responsible department and employee.",
    "Monitor validity and expiry.",
    "Initiate renewal within the defined lead period.",
    "Verify renewed authorization.",
    "Update CCM with the renewed authorization.",
    "Retain historical records and evidence.",
    "Prevent use of expired/cancelled authorizations.",
    "Maintain audit trail of controlled actions."
  ];
}

function apiE3C1EvidenceView_() {
  const auth = requirePermission_("PERM-AUTH-VIEW");
  if (!auth.allowed) return auth.response;

  const kpis = getAuthorizationDashboardKpis_();

  return {
    success: true,
    title: "E3.C1 – Import / Export License & Authorization Control",
    requirement:
      "Wherever applicable, have satisfactory procedures in place for the handling of licenses and authorizations connected to export/import?",
    systemControlStatus: "Implemented",
    evidenceProcedureApprovalStatus:
      kpis.total > 0 ? "Available / Verify Actual Evidence" : "Pending",
    controls: [
      "Authorization register",
      "Applicability assessment",
      "Responsible ownership",
      "Verification",
      "Validity monitoring",
      "Renewal monitoring",
      "Supporting document control",
      "Audit trail"
    ],
    evidence: [
      "Authorization records",
      "Supporting licenses/authorizations",
      "Verification records",
      "Renewal notifications",
      "Renewal records",
      "Procedure document",
      "Audit trail"
    ],
    kpis: kpis,
    procedureFramework: getE3C1ProcedureFramework_()
  };
}

function ccmExecute(action, params) {

  params = params || {};

  const normalizedAction = String(action || "").trim();

  try {

    switch (normalizedAction) {

      /* =========================
         SESSION / DASHBOARD
      ========================= */

      case "me":
        return apiMe_();

      case "dashboard":
        return apiDashboard_();

      /* =========================
         CERTIFICATIONS
      ========================= */

      case "certificates":
        return apiCertificates_(
          params.status || "",
          params.department || ""
        );

      case "certificate360":
        return apiCertificate360_(
          params.certificateId || ""
        );

      case "createCertificate":
        return createCertificate_(
          params.data || params
        );

      case "updateCertificate":
        return updateCertificate_(
          params.certificateId || params.id || "",
          params.data || {}
        );

      /* =========================
         EMPLOYEES
      ========================= */

      case "employees":
        return apiEmployees_();

      case "employee360":
        return apiEmployee360_(
          params.employeeId || params.id || ""
        );

      /* =========================
         RENEWALS
      ========================= */

      case "renewals":
        return apiRenewals_();

      case "runRenewalCheck": {

        const auth =
          requirePermission_("PERM-RENEWAL-MANAGE");

        if (!auth.allowed) {
          return auth.response;
        }

        return runRenewalCheck();
      }

      /* =========================
         DOCUMENTS
      ========================= */

      case "documents":
        return apiDocuments_(
          params.certificateId || ""
        );

      case "documentDownload":
        return authorizeDocumentDownload(
          params.documentId || params.id || ""
        );

      /* =========================
         REPORTS
      ========================= */

      case "reports":
        return apiReports_();

      case "generateReport":
        return generateCCM12Report(
          params.reportType ||
          params.type ||
          "",
          params.options || {}
        );

      case "setupRenewalRules":
        return setupCCM12RenewalRules();

      case "testRenewalEvaluation":
        return testCCM12RenewalEvaluation();

      /* =========================
         IMPORT / EXPORT AUTHORIZATIONS
      ========================= */

      case "authorizations":
        return apiAuthorizations_(
          params.status || "",
          params.applicability || ""
        );

      case "authorization360":
        return apiAuthorization360_(
          params.authorizationId || params.id || ""
        );

      case "createAuthorization":
        return createAuthorization_(
          params.data || params
        );

      case "updateAuthorization":
        return updateAuthorization_(
          params.authorizationId || params.id || "",
          params.data || {}
        );

      case "verifyAuthorization":
        return verifyAuthorization_(
          params.authorizationId || params.id || "",
          params.status || params.verificationStatus || "",
          params.remarks || ""
        );

      case "authorizationRenewals":
        return apiAuthorizationRenewals_();

      case "setupAuthorizationControl":
        return setupImportExportAuthorizationControl_();

      case "authorizationReport":
        return generateCCM12Report(
          "IMPORT_EXPORT_AUTHORIZATION",
          params.options || {}
        );

      case "e3c1Evidence":
        return apiE3C1EvidenceView_();

      /* =========================
         AUDIT
      ========================= */

      case "audit": {

        const auth =
          requirePermission_("PERM-AUDIT-VIEW");

        if (!auth.allowed) {
          return auth.response;
        }

        const rows =
          getSheetObjects_(
            CONFIG.SHEETS.AUDIT_LOG
          );

        return {
          success: true,
          count: rows.length,
          data: rows
            .slice()
            .reverse()
            .slice(0, 500)
            .map(function (row) {
              return {
                Audit_ID:
                  row.Audit_ID ||
                  row.Log_ID ||
                  "",
                Timestamp:
                  formatDateTime_(
                    row.Timestamp ||
                    row.Date ||
                    row.Created_Date
                  ),
                Email:
                  row.Email ||
                  row.User_Email ||
                  "",
                Role_ID:
                  row.Role_ID ||
                  "",
                Action:
                  row.Action ||
                  "",
                Module:
                  row.Module ||
                  "",
                Record_ID:
                  row.Record_ID ||
                  "",
                Description:
                  row.Description ||
                  "",
                Result:
                  row.Result ||
                  ""
              };
            })
        };
      }

      /* =========================
         SETTINGS
      ========================= */

      case "settings": {

        const auth =
          requirePermission_("PERM-RBAC-MANAGE");

        if (!auth.allowed) {
          return auth.response;
        }

        const rows =
          getSheetObjects_(
            CONFIG.SHEETS.SETTINGS
          );

        return {
          success: true,
          count: rows.length,
          data: rows.map(function (row) {
            return {
              Key:
                row.Key ||
                row.Setting ||
                row.Name ||
                "",
              Value:
                row.Value ||
                "",
              Status:
                row.Status ||
                "CONFIGURED"
            };
          })
        };
      }

      /* =========================
         HEALTH
      ========================= */

      case "health":
        return apiHealth_();

      default:
        return {
          success: false,
          error: "UNKNOWN_ACTION"
        };
    }

  } catch (error) {

    logAudit_(
      getCurrentUserEmail_(),
      "API_ERROR",
      "GATEWAY",
      normalizedAction,
      error.message,
      "ERROR"
    );

    return {
      success: false,
      error: "GATEWAY_ERROR",
      message: error.message
    };
  }
}


/* =========================
   MANUAL GATEWAY TESTS
========================= */

function testCcmExecuteMe() {
  return ccmExecute("me", {});
}

function testCcmExecuteDashboard() {
  return ccmExecute("dashboard", {});
}

function testCcmExecuteCertificates() {
  return ccmExecute("certificates", {});
}

function testCcmExecuteEmployees() {
  return ccmExecute("employees", {});
}

function testCcmExecuteRenewals() {
  return ccmExecute("renewals", {});
}

function testCcmExecuteDocuments() {
  return ccmExecute("documents", {});
}

function testCcmExecuteReports() {
  return ccmExecute("reports", {});
}


/************************************************************
 * CCM 12.0 — MANUAL REPORT TESTS
 ************************************************************/

function testCCM12ManagementReport() {
  return generateCCM12Report(
    "MANAGEMENT_SUMMARY",
    {}
  );
}

function testCCM12ExpiryReport() {
  return generateCCM12Report(
    "EXPIRY_RENEWAL",
    {}
  );
}

function testCCM12AuditReport() {
  return generateCCM12Report(
    "AUDIT_TRAIL",
    {}
  );
}
function testCCM12ManagementReport_Diagnostic() {
  try {
    const result = generateCCM12Report("MANAGEMENT_SUMMARY", {});

    console.log("CCM 12.0 REPORT RESULT:");
    console.log(JSON.stringify(result, null, 2));

    SpreadsheetApp.getActiveSpreadsheet().toast(
      "CCM 12.0 Management Report generated successfully",
      "CCM 12.0",
      10
    );

    return result;

  } catch (error) {

    console.error("CCM 12.0 REPORT ERROR:");
    console.error(error && error.stack ? error.stack : String(error));

    SpreadsheetApp.getActiveSpreadsheet().toast(
      "CCM 12.0 Report FAILED: " + String(error),
      "CCM 12.0 ERROR",
      15
    );

    throw error;
  }
}
function testCCM12RenewalRules_Log() {
  const result = setupCCM12RenewalRules();

  Logger.log(
    "CCM 12.0 RENEWAL RULE SETUP:"
  );

  Logger.log(
    JSON.stringify(result, null, 2)
  );

  return result;
}

