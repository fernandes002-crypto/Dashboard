const CONFIG = {

    /*
     * Use the SAME Google Web Client ID as your current working FM dashboard.
     */
    GOOGLE_CLIENT_ID: "50542963972-75ictdqcblqigj1iq30pobdqk06s1v93.apps.googleusercontent.com",

    WORKBOOK_NAME:
        "Copy of FM daily Tracker",

    MAIN_SHEET_NAME:
        "FM collection tracker",

    INVENTORY_SHEET_NAME:
        "Asset Inventory",

    TRANSACTIONS_SHEET_NAME:
        "Asset Transactions",

    LOCATION:
        "UB11",

    OAUTH_SCOPES: [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.readonly"
    ].join(" ")

};
  