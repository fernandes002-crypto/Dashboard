(() => {

    "use strict";


    const SHEETS_API =
        "https://sheets.googleapis.com/v4/spreadsheets";

    const DRIVE_API =
        "https://www.googleapis.com/drive/v3/files";


    const state = {

        idTokenPayload: null,

        accessToken: null,

        workbookId: null,

        workbookName: null,

        inventory: [],

        transactions: [],

        todayLoads: [],

        clients: []

    };


    const $ =
        id => document.getElementById(id);


    document.addEventListener(
        "DOMContentLoaded",
        () => {

            bindEvents();

            setDefaultTimestamp();

            waitForGoogle();

        }
    );


    function bindEvents() {

        $("grant-access")
            .addEventListener(
                "click",
                requestSheetAccess
            );


        $("refresh")
            .addEventListener(
                "click",
                loadDashboard
            );


        $("sign-out")
            .addEventListener(
                "click",
                signOut
            );


        $("movement-form")
            .addEventListener(
                "submit",
                recordMovement
            );


        $("movement")
            .addEventListener(
                "change",
                handleMovementChange
            );


        $("client-cards")
            .addEventListener(
                "click",
                event => {

                    const card =
                        event.target.closest(
                            "[data-client]"
                        );

                    if (card) {

                        openClientDetails(
                            card.dataset.client
                        );

                    }

                }
            );


        $("close-client-details")
            .addEventListener(
                "click",
                closeClientDetails
            );


        document.addEventListener(
            "click",
            event => {

                if (
                    event.target.matches(
                        "[data-close-client-details]"
                    )
                ) {

                    closeClientDetails();

                }

            }
        );


        $("manage-inventory")
            .addEventListener(
                "click",
                openInventoryManager
            );


        $("close-inventory")
            .addEventListener(
                "click",
                closeInventoryManager
            );


        document.addEventListener(
            "click",
            event => {

                if (
                    event.target.matches(
                        "[data-close-modal]"
                    )
                ) {

                    closeInventoryManager();

                }

            }
        );


        $("add-asset")
            .addEventListener(
                "click",
                addAssetType
            );


        $("manage-inventory-body")
            .addEventListener(
                "click",
                event => {

                    const button =
                        event.target.closest(
                            "button[data-action]"
                        );

                    if (!button) {
                        return;
                    }


                    const rowNumber =
                        Number(button.dataset.row);


                    if (
                        button.dataset.action ===
                        "delete"
                    ) {

                        deleteAssetType(
                            rowNumber
                        );

                    }


                    if (
                        button.dataset.action ===
                        "rename"
                    ) {

                        renameAssetType(
                            rowNumber
                        );

                    }

                }
            );

    }


    function waitForGoogle() {

        let checks = 0;

        const maxChecks = 150;


        const timer =
            setInterval(
                () => {

                    checks++;


                    if (
                        window.google?.accounts?.id &&
                        window.google?.accounts?.oauth2
                    ) {

                        clearInterval(timer);

                        initializeGoogle();

                    }


                    if (
                        checks >= maxChecks
                    ) {

                        clearInterval(timer);

                        setAuthStatus(
                            "Google services could not be loaded. Check your internet connection.",
                            true
                        );

                    }

                },
                100
            );

    }


    function initializeGoogle() {

        if (
            !CONFIG.GOOGLE_CLIENT_ID ||
            CONFIG.GOOGLE_CLIENT_ID.includes(
                "PASTE_YOUR"
            )
        ) {

            setAuthStatus(
                "Add your existing Google Web Client ID to config.js.",
                true
            );

            return;

        }


        google.accounts.id.initialize({

            client_id:
                CONFIG.GOOGLE_CLIENT_ID,

            callback:
                handleCredentialResponse,

            auto_select:
                true,

            cancel_on_tap_outside:
                false

        });


        google.accounts.id.renderButton(

            $("google-signin-button"),

            {

                theme:
                    "outline",

                size:
                    "large",

                text:
                    "signin_with",

                shape:
                    "rectangular",

                width:
                    280

            }

        );


        google.accounts.id.prompt();

    }


    function handleCredentialResponse(
        response
    ) {

        try {

            state.idTokenPayload =
                decodeJwtPayload(
                    response.credential
                );


            const name =
                state.idTokenPayload.name ||
                "Google user";


            const email =
                state.idTokenPayload.email ||
                "";


            const picture =
                state.idTokenPayload.picture ||
                "";


            $("user-name").textContent =
                name;


            $("user-email").textContent =
                email;


            if (picture) {

                $("user-photo").src =
                    picture;

                $("user-photo")
                    .classList
                    .remove("hidden");

            }


            $("google-signin-button")
                .classList
                .add("hidden");


            $("sign-out")
                .classList
                .remove("hidden");


            $("grant-access")
                .classList
                .remove("hidden");


            setAuthStatus(
                "Signed in. Click the button below to connect to Google Sheets.",
                false
            );

        } catch (error) {

            console.error(error);

            setAuthStatus(
                "Google sign-in response could not be read.",
                true
            );

        }

    }


    function requestSheetAccess() {

        if (!state.idTokenPayload) {

            setAuthStatus(
                "Sign in with Google first.",
                true
            );

            return;

        }


        const tokenClient =
            google.accounts.oauth2.initTokenClient({

                client_id:
                    CONFIG.GOOGLE_CLIENT_ID,

                scope:
                    CONFIG.OAUTH_SCOPES,

                callback:
                    async tokenResponse => {

                        if (
                            tokenResponse.error
                        ) {

                            console.error(
                                tokenResponse
                            );

                            setAuthStatus(
                                `Google authorization failed: ${tokenResponse.error}`,
                                true
                            );

                            return;

                        }


                        state.accessToken =
                            tokenResponse.access_token;


                        $("grant-access")
                            .classList
                            .add("hidden");


                        $("login-card")
                            .classList
                            .add("hidden");


                        $("dashboard")
                            .classList
                            .remove("hidden");


                        await loadDashboard();

                    }

            });


        tokenClient.requestAccessToken({

            prompt:
                "consent",

            login_hint:
                state.idTokenPayload.email ||
                undefined

        });

    }


    async function loadDashboard() {

        if (!state.accessToken) {
            return;
        }


        setSyncStatus(
            "Syncing with Google Sheets..."
        );


        try {

            const workbook =
                await findWorkbook(
                    CONFIG.WORKBOOK_NAME
                );


            if (!workbook) {

                throw new Error(
                    `Workbook "${CONFIG.WORKBOOK_NAME}" was not found in your Google Drive.`
                );

            }


            state.workbookId =
                workbook.id;


            state.workbookName =
                workbook.name;


            $("workbook-name")
                .textContent =
                workbook.name;


            const metadata =
                await sheetsGet(
                    `/${encodeURIComponent(
                        state.workbookId
                    )}`
                );


            let sheetTitles =
                (metadata.sheets || [])
                    .map(
                        sheet =>
                            sheet.properties.title
                    );


            const missingSheets = [];


            if (
                !sheetTitles.includes(
                    CONFIG.INVENTORY_SHEET_NAME
                )
            ) {

                missingSheets.push(
                    CONFIG.INVENTORY_SHEET_NAME
                );

            }


            if (
                !sheetTitles.includes(
                    CONFIG.TRANSACTIONS_SHEET_NAME
                )
            ) {

                missingSheets.push(
                    CONFIG.TRANSACTIONS_SHEET_NAME
                );

            }


            if (missingSheets.length) {

                await createSheets(
                    missingSheets
                );

                sheetTitles =
                    sheetTitles.concat(
                        missingSheets
                    );

            }


            const [

                inventoryRows,

                transactionRows,

                mainRows

            ] = await Promise.all([

                getValues(
                    CONFIG.INVENTORY_SHEET_NAME
                ),

                getValues(
                    CONFIG.TRANSACTIONS_SHEET_NAME
                ),

                getValues(
                    CONFIG.MAIN_SHEET_NAME
                )

            ]);


            state.inventory =
                parseInventory(
                    inventoryRows
                );


            state.transactions =
                parseTransactions(
                    transactionRows
                );


            state.todayLoads =
                parseTodayLoads(
                    mainRows
                );


            state.clients =
                parseClients(
                    mainRows
                );


            renderDashboard();


            setSyncStatus(
                `Synced at ${new Date().toLocaleTimeString()}`
            );


        } catch (error) {

            console.error(error);

            setSyncStatus(
                error.message ||
                "Unable to load spreadsheet.",
                true
            );

        }

    }


    async function findWorkbook(name) {

        const query = [

            `name = '${escapeDriveQuery(name)}'`,

            `mimeType = 'application/vnd.google-apps.spreadsheet'`,

            `trashed = false`

        ].join(" and ");


        const url =

            `${DRIVE_API}?q=${encodeURIComponent(query)}` +

            `&pageSize=20` +

            `&fields=files(id,name,mimeType,modifiedTime,webViewLink)`;


        const data =
            await fetchJson(
                url,
                {
                    headers:
                        authHeaders()
                }
            );


        return data.files?.[0] ||
            null;

    }


    async function createSheets(
        names
    ) {

        const requests =
            names.map(
                title => ({

                    addSheet: {

                        properties: {

                            title

                        }

                    }

                })
            );


        await sheetsPost(

            `/${encodeURIComponent(
                state.workbookId
            )}:batchUpdate`,

            {
                requests
            }

        );


        if (
            names.includes(
                CONFIG.INVENTORY_SHEET_NAME
            )
        ) {

            await updateValues(

                CONFIG.INVENTORY_SHEET_NAME,

                [
                    [
                        "Asset",
                        "Balance"
                    ]
                ]

            );

        }


        if (
            names.includes(
                CONFIG.TRANSACTIONS_SHEET_NAME
            )
        ) {

            await updateValues(

                CONFIG.TRANSACTIONS_SHEET_NAME,

                [
                    [
                        "Timestamp",
                        "Client",
                        "Movement",
                        "Asset",
                        "Quantity",
                        "User"
                    ]
                ]

            );

        }

    }


    async function getValues(
        sheetName
    ) {

        const range =
            `${quoteSheetName(
                sheetName
            )}!A:AE`;


        const data =
            await sheetsGet(

                `/${encodeURIComponent(
                    state.workbookId
                )}/values/${encodeURIComponent(
                    range
                )}`

            );


        return data.values || [];

    }


    async function updateValues(
        sheetName,
        rows
    ) {

        const range =
            `${quoteSheetName(
                sheetName
            )}!A1`;


        return sheetsPut(

            `/${encodeURIComponent(
                state.workbookId
            )}/values/${encodeURIComponent(
                range
            )}?valueInputOption=USER_ENTERED`,

            {

                range,

                majorDimension:
                    "ROWS",

                values:
                    rows

            }

        );

    }


    function parseInventory(
        rows
    ) {

        if (!rows.length) {
            return [];
        }


        const header =
            rows[0].map(
                normalizeHeader
            );


        const assetIdx =
            findColumn(

                header,

                [
                    "asset",
                    "asset name",
                    "item",
                    "type"
                ]

            );


        const balanceIdx =
            findColumn(

                header,

                [
                    "balance",
                    "current balance",
                    "stock",
                    "quantity"
                ]

            );


        if (assetIdx < 0) {
            return [];
        }


        return rows

            .slice(1)

            .map(
                (row, index) => ({

                    rowNumber:
                        index + 2,

                    asset:
                        String(
                            row[assetIdx] ?? ""
                        ).trim(),

                    balance:
                        balanceIdx >= 0
                            ? numericValue(
                                row[balanceIdx]
                            )
                            : 0,

                    assetColumn:
                        assetIdx + 1,

                    balanceColumn:
                        balanceIdx >= 0
                            ? balanceIdx + 1
                            : 2

                })
            )

            .filter(
                item =>
                    item.asset
            );

    }


    function parseTransactions(
        rows
    ) {

        if (!rows.length) {
            return [];
        }


        const header =
            rows[0].map(
                normalizeHeader
            );


        const idx = {

            timestamp:
                findColumn(
                    header,
                    [
                        "timestamp",
                        "date",
                        "datetime"
                    ]
                ),

            client:
                findColumn(
                    header,
                    [
                        "client",
                        "client name"
                    ]
                ),

            movement:
                findColumn(
                    header,
                    [
                        "movement",
                        "type",
                        "direction"
                    ]
                ),

            asset:
                findColumn(
                    header,
                    [
                        "asset",
                        "asset name",
                        "item"
                    ]
                ),

            quantity:
                findColumn(
                    header,
                    [
                        "quantity",
                        "qty"
                    ]
                ),

            user:
                findColumn(
                    header,
                    [
                        "user",
                        "entered by",
                        "email"
                    ]
                )

        };


        return rows

            .slice(1)

            .map(
                row => ({

                    timestamp:
                        idx.timestamp >= 0
                            ? row[idx.timestamp] ?? ""
                            : "",

                    client:
                        idx.client >= 0
                            ? row[idx.client] ?? ""
                            : "",

                    movement:
                        idx.movement >= 0
                            ? row[idx.movement] ?? ""
                            : "",

                    asset:
                        idx.asset >= 0
                            ? row[idx.asset] ?? ""
                            : "",

                    quantity:
                        idx.quantity >= 0
                            ? numericValue(
                                row[idx.quantity]
                            )
                            : 0,

                    user:
                        idx.user >= 0
                            ? row[idx.user] ?? ""
                            : ""

                })
            )

            .filter(
                item =>
                    item.asset ||
                    item.client
            );

    }


    function findMainHeaderRow(
        rows
    ) {

        return rows.findIndex(
            row => {

                const headers =
                    row.map(
                        normalizeHeader
                    );


                const hasClient =
                    headers.includes("cleint") ||
                    headers.includes("client") ||
                    headers.includes("collection client");


                const hasLoadType =
                    headers.includes("load type") ||
                    headers.includes("loadtype");


                const hasArrival =
                    headers.includes("planned arrival") ||
                    headers.includes("plannedarrival");


                return (
                    hasClient &&
                    hasLoadType &&
                    hasArrival
                );

            }
        );

    }


    function parseClients(
        rows
    ) {

        const headerRowIndex =
            findMainHeaderRow(
                rows
            );


        if (headerRowIndex < 0) {
            return [];
        }


        const headers =
            rows[
                headerRowIndex
            ].map(
                normalizeHeader
            );


        const clientIdx =
            findColumn(

                headers,

                [
                    "cleint",
                    "client",
                    "collection client"
                ]

            );


        if (clientIdx < 0) {
            return [];
        }


        const clients =
            new Set();


        for (
            let i =
                headerRowIndex + 1;

            i < rows.length;

            i++
        ) {

            const value =
                String(
                    rows[i][clientIdx] ??
                    ""
                ).trim();


            if (value) {
                clients.add(value);
            }

        }


        return [
            ...clients
        ].sort(
            (a, b) =>
                a.localeCompare(b)
        );

    }


    function parseTodayLoads(
        rows
    ) {

        if (!rows.length) {
            return [];
        }


        const headerRowIndex =
            findMainHeaderRow(
                rows
            );


        if (headerRowIndex < 0) {
            return [];
        }


        const headers =
            rows[
                headerRowIndex
            ].map(
                normalizeHeader
            );


        const clientIdx =
            findColumn(
                headers,
                [
                    "cleint",
                    "client",
                    "collection client"
                ]
            );


        const loadTypeIdx =
            findColumn(
                headers,
                [
                    "load type",
                    "loadtype"
                ]
            );


        const plannedArrivalIdx =
            findColumn(
                headers,
                [
                    "planned arrival",
                    "plannedarrival"
                ]
            );


        const palletsIdx =
            findColumn(
                headers,
                [
                    "pallets",
                    "pallet"
                ]
            );


        const looseIdx =
            findColumn(
                headers,
                [
                    "bags / loose parcles",
                    "bags / loose parcels",
                    "bags loose parcles",
                    "bags loose parcels"
                ]
            );


        const locationIdx =
            findColumn(
                headers,
                [
                    "/",
                    "location",
                    "site"
                ]
            );


        if (
            clientIdx < 0 ||
            loadTypeIdx < 0 ||
            plannedArrivalIdx < 0
        ) {

            return [];

        }


        const today =
            new Date();


        const todayDay =
            today.getDate();


        const todayMonth =
            today.getMonth() + 1;


        const todayYear =
            today.getFullYear();


        let activeDate =
            null;


        const loads = [];


        for (
            let i =
                headerRowIndex + 1;

            i < rows.length;

            i++
        ) {

            const row =
                rows[i] || [];


            let sectionDate =
                null;


            for (
                let col = 0;

                col < Math.min(
                    4,
                    row.length
                );

                col++
            ) {

                sectionDate =
                    parseTrackerDate(
                        row[col]
                    );


                if (sectionDate) {
                    break;
                }

            }


            if (sectionDate) {

                activeDate =
                    sectionDate;

                continue;

            }


            const client =
                String(
                    row[clientIdx] ??
                    ""
                ).trim();


            const loadType =
                String(
                    row[loadTypeIdx] ??
                    ""
                ).trim();


            const plannedArrival =
                String(
                    row[plannedArrivalIdx] ??
                    ""
                ).trim();


            const location =
                locationIdx >= 0
                    ? String(
                        row[locationIdx] ??
                        ""
                    ).trim()
                    : "";


            const pallets =
                palletsIdx >= 0
                    ? String(
                        row[palletsIdx] ??
                        ""
                    ).trim()
                    : "";


            const loose =
                looseIdx >= 0
                    ? String(
                        row[looseIdx] ??
                        ""
                    ).trim()
                    : "";


            if (
                !client &&
                !loadType &&
                !plannedArrival
            ) {
                continue;
            }


            if (
                client.toLowerCase() ===
                    "cleint" ||

                client.toLowerCase() ===
                    "client"
            ) {
                continue;
            }


            if (
                location &&
                location.toUpperCase() !==
                    CONFIG.LOCATION.toUpperCase()
            ) {
                continue;
            }


            if (!activeDate) {
                continue;
            }


            if (
                activeDate.day !==
                    todayDay ||

                activeDate.month !==
                    todayMonth ||

                activeDate.year !==
                    todayYear
            ) {
                continue;
            }


            const direction =
                loadType.toUpperCase() ===
                    "RETURNS"

                    ? "Outbound"

                    : "Inbound";


            loads.push({

                client,

                loadType,

                plannedArrival,

                pallets,

                loose,

                direction

            });

        }


        return loads;

    }


    function parseTrackerDate(
        value
    ) {

        if (
            value === null ||
            value === undefined
        ) {
            return null;
        }


        const text =
            String(value).trim();


        if (!text) {
            return null;
        }


        const shortDate =
            text.match(
                /^(\d{1,2})[\/\-](\d{1,2})$/
            );


        if (shortDate) {

            return {

                day:
                    Number(
                        shortDate[1]
                    ),

                month:
                    Number(
                        shortDate[2]
                    ),

                year:
                    new Date()
                        .getFullYear()

            };

        }


        const parsed =
            new Date(text);


        if (
            !Number.isNaN(
                parsed.getTime()
            )
        ) {

            return {

                day:
                    parsed.getDate(),

                month:
                    parsed.getMonth() + 1,

                year:
                    parsed.getFullYear()

            };

        }


        return null;

    }


    /* =====================================================
       DASHBOARD
       ===================================================== */


    function renderDashboard() {

        const inbound =
            state.todayLoads.filter(
                item =>
                    item.direction ===
                    "Inbound"
            );


        const outbound =
            state.todayLoads.filter(
                item =>
                    item.direction ===
                    "Outbound"
            );


        $("inbound-badge")
            .textContent =
            inbound.length
                .toLocaleString();


        $("outbound-badge")
            .textContent =
            outbound.length
                .toLocaleString();


        /*
         * NEW:
         * Warehouse asset cards
         */

        renderWarehouseAssets();


        renderClientCards();


        renderLoadTable(
            $("inbound-body"),
            inbound
        );


        renderLoadTable(
            $("outbound-body"),
            outbound
        );


        /*
         * Asset dropdown
         */

        $("asset").innerHTML =
            state.inventory.length

                ? `

                    <option value="">
                        Select asset
                    </option>

                    ${state.inventory.map(
                        item => `

                            <option
                                value="${escapeAttr(
                                    item.asset
                                )}"
                            >
                                ${escapeHtml(
                                    item.asset
                                )}
                            </option>

                        `
                    ).join("")}

                  `

                : `
                    <option value="">
                        No assets configured
                    </option>
                  `;


        /*
         * Client dropdown
         */

        $("client").innerHTML =
            state.clients.length

                ? `

                    <option value="">
                        Select client
                    </option>

                    ${state.clients.map(
                        client => `

                            <option
                                value="${escapeAttr(
                                    client
                                )}"
                            >
                                ${escapeHtml(
                                    client
                                )}
                            </option>

                        `
                    ).join("")}

                  `

                : `
                    <option value="">
                        No clients found
                    </option>
                  `;


        /*
         * Recent transactions
         */

        $("transactions-body").innerHTML =
            state.transactions.length

                ? state.transactions

                    .slice(-20)

                    .reverse()

                    .map(
                        item => `

                            <tr>

                                <td>
                                    ${escapeHtml(
                                        String(
                                            item.timestamp
                                        )
                                    )}
                                </td>

                                <td>
                                    ${escapeHtml(
                                        String(
                                            item.client
                                        )
                                    )}
                                </td>

                                <td>

                                    <span
                                        class="movement-badge ${movementClass(
                                            item.movement
                                        )}"
                                    >
                                        ${escapeHtml(
                                            String(
                                                item.movement
                                            )
                                        )}
                                    </span>

                                </td>

                                <td>
                                    ${escapeHtml(
                                        String(
                                            item.asset
                                        )
                                    )}
                                </td>

                                <td class="num">
                                    ${formatNumber(
                                        item.quantity
                                    )}
                                </td>

                                <td>
                                    ${escapeHtml(
                                        String(
                                            item.user
                                        )
                                    )}
                                </td>

                            </tr>

                        `
                    )

                    .join("")

                : emptyRow(
                    6,
                    "No asset movements recorded yet."
                );


        $("dashboard-date")
            .textContent =
            new Date().toLocaleDateString(
                "en-GB",
                {

                    weekday:
                        "long",

                    day:
                        "numeric",

                    month:
                        "long",

                    year:
                        "numeric"

                }
            );


        renderInventoryManager();

    }


    /*
     * =====================================================
     * NEW WAREHOUSE ASSET CARDS
     * =====================================================
     *
     * Current number:
     *     state.inventory balance
     *
     * Movement indicator:
     *
     * RECEIVED = +
     * SENT     = -
     * DISCARD  = -
     *
     * Only today's transactions are included.
     */


    function getWarehouseAssetMovements() {

        const movements = {};


        state.transactions

            .filter(
                item =>
                    isTodayTransaction(
                        item.timestamp
                    )
            )

            .forEach(
                item => {

                    const asset =
                        String(
                            item.asset ?? ""
                        ).trim();


                    if (!asset) {
                        return;
                    }


                    const movement =
                        String(
                            item.movement ?? ""
                        )
                        .trim()
                        .toUpperCase();


                    const quantity =
                        Number(
                            item.quantity
                        ) || 0;


                    if (
                        !movements[
                            asset.toLowerCase()
                        ]
                    ) {

                        movements[
                            asset.toLowerCase()
                        ] = {

                            asset,

                            net:
                                0,

                            received:
                                0,

                            sent:
                                0,

                            discarded:
                                0

                        };

                    }


                    const entry =
                        movements[
                            asset.toLowerCase()
                        ];


                    if (
                        movement ===
                        "RECEIVED"
                    ) {

                        entry.net +=
                            quantity;

                        entry.received +=
                            quantity;

                    }


                    if (
                        movement ===
                        "SENT"
                    ) {

                        entry.net -=
                            quantity;

                        entry.sent +=
                            quantity;

                    }


                    if (
                        movement ===
                        "DISCARD"
                    ) {

                        entry.net -=
                            quantity;

                        entry.discarded +=
                            quantity;

                    }

                }
            );


        return movements;

    }


    function renderWarehouseAssets() {
        const container = $("warehouse-assets");
    
        if (!state.inventory.length) {
            container.innerHTML = `
                <div class="warehouse-empty">
                    No assets configured.
                </div>
            `;
            return;
        }
    
        container.innerHTML = state.inventory.map(item => {
    
            const assetName = String(item.asset).trim();
            const balance = Number(item.balance) || 0;
    
            /*
             * Calculate today's movement for this asset.
             *
             * RECEIVED = positive
             * SENT     = negative
             * DISCARD  = negative
             */
            const movementTotal = state.transactions
                .filter(transaction => {
                    return (
                        String(transaction.asset).trim().toLowerCase() ===
                        assetName.toLowerCase()
                        &&
                        isTodayTransaction(transaction.timestamp)
                    );
                })
                .reduce((total, transaction) => {
    
                    const movement =
                        String(transaction.movement)
                            .trim()
                            .toUpperCase();
    
                    const quantity =
                        Number(transaction.quantity) || 0;
    
                    if (movement === "RECEIVED") {
                        return total + quantity;
                    }
    
                    if (
                        movement === "SENT" ||
                        movement === "DISCARD"
                    ) {
                        return total - quantity;
                    }
    
                    return total;
                }, 0);
    
            let indicator = "";
    
            if (movementTotal > 0) {
                indicator = `
                    <span class="asset-movement positive">
                        +${formatNumber(movementTotal)}
                    </span>
                `;
            } else if (movementTotal < 0) {
                indicator = `
                    <span class="asset-movement negative">
                        ${formatNumber(movementTotal)}
                    </span>
                `;
            } else {
                indicator = `
                    <span class="asset-movement neutral">
                        0
                    </span>
                `;
            }
    
            return `
                <div class="warehouse-asset-card">
    
                    <div class="warehouse-asset-name">
                        ${escapeHtml(assetName)}
                    </div>
    
                    <div class="warehouse-asset-count">
                        ${formatNumber(balance)}
                    </div>
    
                    <div class="warehouse-asset-movement">
                        ${indicator}
                        <span class="movement-label">today</span>
                    </div>
    
                </div>
            `;
        }).join("");
    }


    function getClientSummary(
        client
    ) {

        const loads =
            state.todayLoads.filter(
                item =>
                    item.client.toLowerCase() ===
                    client.toLowerCase()
            );


        const transactions =
            state.transactions.filter(
                item => {

                    if (
                        String(
                            item.client
                        )
                            .trim()
                            .toLowerCase() !==
                        client.toLowerCase()
                    ) {

                        return false;

                    }


                    return isTodayTransaction(
                        item.timestamp
                    );

                }
            );


        const inboundLoads =
            loads.filter(
                item =>
                    item.direction ===
                    "Inbound"
            );


        const outboundLoads =
            loads.filter(
                item =>
                    item.direction ===
                    "Outbound"
            );


        const received =
            transactions.filter(
                item =>
                    String(
                        item.movement
                    )
                        .trim()
                        .toUpperCase() ===
                    "RECEIVED"
            );


        const sent =
            transactions.filter(
                item =>
                    String(
                        item.movement
                    )
                        .trim()
                        .toUpperCase() ===
                    "SENT"
            );


        const discarded =
            transactions.filter(
                item =>
                    String(
                        item.movement
                    )
                        .trim()
                        .toUpperCase() ===
                    "DISCARD"
            );


        const expectedPallets =
            sumNumericText(
                inboundLoads,
                "pallets"
            );


        const expectedLoose =
            sumNumericText(
                inboundLoads,
                "loose"
            );


        const outboundPallets =
            sumNumericText(
                outboundLoads,
                "pallets"
            );


        const outboundLoose =
            sumNumericText(
                outboundLoads,
                "loose"
            );


        const actualReceived =
            received.reduce(
                (
                    sum,
                    item
                ) =>
                    sum +
                    item.quantity,
                0
            );


        const actualSent =
            sent.reduce(
                (
                    sum,
                    item
                ) =>
                    sum +
                    item.quantity,
                0
            );


        const actualDiscarded =
            discarded.reduce(
                (
                    sum,
                    item
                ) =>
                    sum +
                    item.quantity,
                0
            );


        return {

            client,

            loads,

            inboundLoads,

            outboundLoads,

            transactions,

            received,

            sent,

            discarded,

            expectedPallets,

            expectedLoose,

            outboundPallets,

            outboundLoose,

            actualReceived,

            actualSent,

            actualDiscarded,

            expectedInboundTotal:
                expectedPallets +
                expectedLoose,

            expectedOutboundTotal:
                outboundPallets +
                outboundLoose,

            actualInboundTotal:
                actualReceived,

            actualOutboundTotal:
                actualSent +
                actualDiscarded

        };

    }


    function renderClientCards() {

        const container =
            $("client-cards");


        const clients =
            new Set([

                ...state.clients,

                ...state.todayLoads
                    .map(
                        item =>
                            item.client
                    )
                    .filter(Boolean),

                ...state.transactions

                    .filter(
                        item =>
                            isTodayTransaction(
                                item.timestamp
                            )
                    )

                    .map(
                        item =>
                            String(
                                item.client
                            ).trim()
                    )

                    .filter(Boolean)

            ]);


        const summaries =
            [...clients]

                .filter(Boolean)

                .sort(
                    (a, b) =>
                        a.localeCompare(b)
                )

                .map(
                    getClientSummary
                );


        container.innerHTML =
            summaries.length

                ? summaries.map(
                    summary => `

                        <article
                            class="client-card ${getClientCardDirectionClass(
                                summary
                            )}"
                            data-client="${escapeAttr(
                                summary.client
                            )}"
                            tabindex="0"
                            role="button"
                        >

                            <div
                                class="client-card-top"
                            >

                                <div>

                                    <div class="eyebrow">
                                        CLIENT
                                    </div>

                                    <h3>
                                        ${escapeHtml(
                                            summary.client
                                        )}
                                    </h3>

                                </div>

                                <span
                                    class="client-card-arrow"
                                >
                                    →
                                </span>

                            </div>


                            <div class="direction-row">

                                ${
                                    summary.inboundLoads.length

                                        ? `

                                            <span
                                                class="direction-pill inbound-pill"
                                            >
                                                <span
                                                    class="direction-dot"
                                                ></span>

                                                Inbound
                                            </span>

                                          `

                                        : ""
                                }


                                ${
                                    summary.outboundLoads.length

                                        ? `

                                            <span
                                                class="direction-pill outbound-pill"
                                            >
                                                <span
                                                    class="direction-dot"
                                                ></span>

                                                Outbound
                                            </span>

                                          `

                                        : ""
                                }


                                ${
                                    !summary.inboundLoads.length &&
                                    !summary.outboundLoads.length

                                        ? `

                                            <span
                                                class="direction-pill neutral-pill"
                                            >
                                                Asset movements
                                            </span>

                                          `

                                        : ""
                                }

                            </div>


                            <div class="client-metrics">

                                <div class="client-metric">

                                    <span>
                                        Expected
                                    </span>

                                    <strong>
                                        ${formatNumber(
                                            summary.expectedInboundTotal
                                        )}
                                    </strong>

                                    <small>
                                        ${formatNumber(
                                            summary.expectedPallets
                                        )}
                                        pallets ·
                                        ${formatNumber(
                                            summary.expectedLoose
                                        )}
                                        loose
                                    </small>

                                </div>


                                <div
                                    class="client-metric actual-metric"
                                >

                                    <span>
                                        Actual
                                    </span>

                                    <strong>
                                        ${formatNumber(
                                            summary.actualInboundTotal
                                        )}
                                    </strong>

                                    <small>
                                        ${formatNumber(
                                            summary.received.length
                                        )}
                                        received movement${summary.received.length === 1 ? "" : "s"}
                                    </small>

                                </div>

                            </div>


                            <div
                                class="client-card-footer"
                            >

                                <span>

                                    ${summary.transactions.length}

                                    transaction${summary.transactions.length === 1 ? "" : "s"}
                                    today

                                </span>

                                <span>
                                    View details
                                </span>

                            </div>

                        </article>

                    `
                ).join("")

                : `

                    <div class="client-cards-empty">

                        <strong>
                            No client activity for today
                        </strong>

                        <span>
                            Clients will appear here when today's loads or asset movements are available.
                        </span>

                    </div>

                  `;

    }


    function getClientCardDirectionClass(
        summary
    ) {

        if (
            summary.inboundLoads.length &&
            summary.outboundLoads.length
        ) {

            return "direction-mixed";

        }


        if (
            summary.outboundLoads.length
        ) {

            return "direction-outbound";

        }


        if (
            summary.inboundLoads.length
        ) {

            return "direction-inbound";

        }


        return "direction-neutral";

    }


    function sumNumericText(
        items,
        key
    ) {

        return items.reduce(
            (
                sum,
                item
            ) => {

                const value =
                    String(
                        item[key] ?? ""
                    )
                        .replace(
                            /,/g,
                            ""
                        )
                        .trim();


                const parsed =
                    Number(value);


                return Number.isFinite(
                    parsed
                )
                    ? sum + parsed
                    : sum;

            },
            0
        );

    }


    function isTodayTransaction(
        timestamp
    ) {

        if (!timestamp) {
            return false;
        }


        const parsed =
            new Date(timestamp);


        if (
            Number.isNaN(
                parsed.getTime()
            )
        ) {

            return (
                String(
                    timestamp
                ).slice(0, 10) ===
                new Date()
                    .toISOString()
                    .slice(0, 10)
            );

        }


        const now =
            new Date();


        return (

            parsed.getFullYear() ===
                now.getFullYear() &&

            parsed.getMonth() ===
                now.getMonth() &&

            parsed.getDate() ===
                now.getDate()

        );

    }


    function openClientDetails(
        client
    ) {

        const summary =
            getClientSummary(
                client
            );


        $("client-details-title")
            .textContent =
            summary.client;


        $("client-details-subtitle")
            .textContent =
            `${summary.loads.length} load${summary.loads.length === 1 ? "" : "s"} · ` +
            `${summary.transactions.length} asset transaction${summary.transactions.length === 1 ? "" : "s"} today`;


        $("client-details-summary")
            .innerHTML = `

                <div
                    class="detail-stat inbound-detail"
                >

                    <span>
                        Inbound expected
                    </span>

                    <strong>
                        ${formatNumber(
                            summary.expectedInboundTotal
                        )}
                    </strong>

                    <small>
                        ${formatNumber(
                            summary.expectedPallets
                        )}
                        pallets ·
                        ${formatNumber(
                            summary.expectedLoose
                        )}
                        loose
                    </small>

                </div>


                <div
                    class="detail-stat inbound-detail"
                >

                    <span>
                        Inbound actual
                    </span>

                    <strong>
                        ${formatNumber(
                            summary.actualInboundTotal
                        )}
                    </strong>

                    <small>
                        ${formatNumber(
                            summary.received.length
                        )}
                        received movement${summary.received.length === 1 ? "" : "s"}
                    </small>

                </div>


                <div
                    class="detail-stat outbound-detail"
                >

                    <span>
                        Outbound expected
                    </span>

                    <strong>
                        ${formatNumber(
                            summary.expectedOutboundTotal
                        )}
                    </strong>

                    <small>
                        ${formatNumber(
                            summary.outboundPallets
                        )}
                        pallets ·
                        ${formatNumber(
                            summary.outboundLoose
                        )}
                        loose
                    </small>

                </div>


                <div
                    class="detail-stat outbound-detail"
                >

                    <span>
                        Outbound actual
                    </span>

                    <strong>
                        ${formatNumber(
                            summary.actualOutboundTotal
                        )}
                    </strong>

                    <small>
                        ${formatNumber(
                            summary.sent.length
                        )}
                        sent ·
                        ${formatNumber(
                            summary.discarded.length
                        )}
                        discarded
                    </small>

                </div>

            `;


        $("client-details-loads")
            .innerHTML =

            summary.loads.length

                ? summary.loads.map(
                    load => `

                        <tr>

                            <td>

                                <strong>
                                    ${escapeHtml(
                                        load.direction
                                    )}
                                </strong>

                                <span
                                    class="subtext"
                                >
                                    ${escapeHtml(
                                        load.loadType ||
                                        ""
                                    )}
                                </span>

                            </td>

                            <td>
                                ${escapeHtml(
                                    load.plannedArrival ||
                                    "Not planned"
                                )}
                            </td>

                            <td class="num">
                                ${escapeHtml(
                                    load.pallets ||
                                    "0"
                                )}
                            </td>

                            <td>
                                ${escapeHtml(
                                    load.loose ||
                                    "0"
                                )}
                            </td>

                        </tr>

                    `
                ).join("")

                : emptyRow(
                    4,
                    "No loads planned for this client today."
                );


        $("client-details-transactions")
            .innerHTML =

            summary.transactions.length

                ? summary.transactions

                    .slice()

                    .sort(
                        (a, b) =>
                            new Date(
                                b.timestamp
                            ) -
                            new Date(
                                a.timestamp
                            )
                    )

                    .map(
                        item => `

                            <tr>

                                <td>
                                    ${formatTransactionTime(
                                        item.timestamp
                                    )}
                                </td>

                                <td>

                                    <span
                                        class="movement-badge ${movementClass(
                                            item.movement
                                        )}"
                                    >
                                        ${escapeHtml(
                                            String(
                                                item.movement
                                            )
                                        )}
                                    </span>

                                </td>

                                <td>

                                    <strong>
                                        ${escapeHtml(
                                            String(
                                                item.asset
                                            )
                                        )}
                                    </strong>

                                </td>

                                <td class="num">
                                    ${formatNumber(
                                        item.quantity
                                    )}
                                </td>

                                <td>
                                    ${escapeHtml(
                                        String(
                                            item.user ||
                                            ""
                                        )
                                    )}
                                </td>

                            </tr>

                        `
                    )

                    .join("")

                : emptyRow(
                    5,
                    "No asset transactions recorded for this client today."
                );


        $("client-details-modal")
            .classList
            .remove("hidden");

    }


    function closeClientDetails() {

        $("client-details-modal")
            .classList
            .add("hidden");

    }


    function formatTransactionTime(
        timestamp
    ) {

        const parsed =
            new Date(timestamp);


        if (
            Number.isNaN(
                parsed.getTime()
            )
        ) {

            return escapeHtml(
                String(
                    timestamp || ""
                )
            );

        }


        return escapeHtml(

            parsed.toLocaleTimeString(
                "en-GB",
                {

                    hour:
                        "2-digit",

                    minute:
                        "2-digit"

                }
            )

        );

    }


    function movementClass(
        movement
    ) {

        const value =
            String(
                movement
            )
                .trim()
                .toUpperCase();


        if (
            value ===
            "RECEIVED"
        ) {

            return "movement-received";

        }


        if (
            value ===
            "SENT"
        ) {

            return "movement-sent";

        }


        if (
            value ===
            "DISCARD"
        ) {

            return "movement-discard";

        }


        return "movement-other";

    }


    function renderLoadTable(
        body,
        items
    ) {

        body.innerHTML =

            items.length

                ? items.map(
                    item => `

                        <tr>

                            <td>

                                <strong>
                                    ${escapeHtml(
                                        item.client ||
                                        "Unknown"
                                    )}
                                </strong>

                                ${
                                    item.loadType

                                        ? `

                                            <span
                                                class="subtext"
                                            >
                                                ${escapeHtml(
                                                    item.loadType
                                                )}
                                            </span>

                                          `

                                        : ""
                                }

                            </td>

                            <td>
                                ${escapeHtml(
                                    item.plannedArrival ||
                                    "Not planned"
                                )}
                            </td>

                            <td class="num">
                                ${escapeHtml(
                                    item.pallets ||
                                    "0"
                                )}
                            </td>

                            <td>
                                ${escapeHtml(
                                    item.loose ||
                                    "0"
                                )}
                            </td>

                        </tr>

                    `
                ).join("")

                : emptyRow(
                    4,
                    "No matching loads found for today."
                );

    }


    async function recordMovement(
        event
    ) {

        event.preventDefault();


        const movement =
            $("movement").value;


        /*
         * Discard is always warehouse stock.
         * Keep the client value consistent.
         */

        const client =
            movement === "DISCARD"

                ? "HSC London (Self)"

                : (
                    $("client").value.trim() ||
                    "HSC London (Self)"
                );


        const asset =
            $("asset").value;


        const quantity =
            Number(
                $("quantity").value
            );


        const user =
            state.idTokenPayload?.email ||
            state.idTokenPayload?.name ||
            "Google user";


        if (
            !client ||
            !asset ||
            !Number.isInteger(
                quantity
            ) ||
            quantity <= 0
        ) {

            setMovementStatus(
                "Select a client, asset and enter a whole quantity greater than zero.",
                true
            );

            return;

        }


        const item =
            state.inventory.find(
                inventoryItem =>
                    inventoryItem.asset
                        .toLowerCase() ===
                    asset.toLowerCase()
            );


        if (!item) {

            setMovementStatus(
                "That asset is not present in Inventory.",
                true
            );

            return;

        }


        if (
            movement !== "RECEIVED" &&
            quantity > item.balance
        ) {

            setMovementStatus(
                `Cannot remove ${quantity}. Current ${item.asset} balance is ${item.balance}.`,
                true
            );

            return;

        }


        const newBalance =
            movement === "RECEIVED"

                ? item.balance +
                    quantity

                : item.balance -
                    quantity;


        try {

            setMovementStatus(
                "Recording movement..."
            );


            const transactionRange =
                `${quoteSheetName(
                    CONFIG.TRANSACTIONS_SHEET_NAME
                )}!A:F`;


            await sheetsPost(

                `/${encodeURIComponent(
                    state.workbookId
                )}/values/${encodeURIComponent(
                    transactionRange
                )}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,

                {

                    values: [[

                        new Date().toISOString(),

                        client,

                        movement,

                        asset,

                        quantity,

                        user

                    ]]

                }

            );


            const balanceCell =
                columnLetter(
                    item.balanceColumn
                ) +
                item.rowNumber;


            const inventoryRange =
                `${quoteSheetName(
                    CONFIG.INVENTORY_SHEET_NAME
                )}!${balanceCell}`;


            await sheetsPut(

                `/${encodeURIComponent(
                    state.workbookId
                )}/values/${encodeURIComponent(
                    inventoryRange
                )}?valueInputOption=USER_ENTERED`,

                {

                    range:
                        inventoryRange,

                    majorDimension:
                        "ROWS",

                    values: [
                        [
                            newBalance
                        ]
                    ]

                }

            );


            $("movement-form")
                .reset();


            setDefaultTimestamp();


            setMovementStatus(
                "Movement recorded successfully."
            );


            await loadDashboard();


        } catch (error) {

            console.error(error);


            setMovementStatus(
                error.message,
                true
            );

        }

    }


    async function addAssetType() {

        const input =
            $("new-asset-name");


        const balanceInput =
            $("new-asset-balance");


        const assetName =
            input.value.trim();


        const balance =
            Number(
                balanceInput.value
            );


        if (!assetName) {

            setInventoryManageStatus(
                "Enter an asset type name.",
                true
            );

            return;

        }


        if (
            !Number.isInteger(
                balance
            ) ||
            balance < 0
        ) {

            setInventoryManageStatus(
                "Starting balance must be a whole number of zero or more.",
                true
            );

            return;

        }


        const exists =
            state.inventory.some(
                item =>
                    item.asset
                        .toLowerCase() ===
                    assetName.toLowerCase()
            );


        if (exists) {

            setInventoryManageStatus(
                "That asset type already exists.",
                true
            );

            return;

        }


        try {

            setInventoryManageStatus(
                "Adding asset type..."
            );


            const range =
                `${quoteSheetName(
                    CONFIG.INVENTORY_SHEET_NAME
                )}!A:B`;


            await sheetsPost(

                `/${encodeURIComponent(
                    state.workbookId
                )}/values/${encodeURIComponent(
                    range
                )}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,

                {

                    values: [[

                        assetName,

                        balance

                    ]]

                }

            );


            input.value = "";

            balanceInput.value =
                "0";


            setInventoryManageStatus(
                "Asset type added successfully."
            );


            await loadDashboard();


        } catch (error) {

            console.error(error);


            setInventoryManageStatus(
                error.message,
                true
            );

        }

    }


    async function deleteAssetType(
        rowNumber
    ) {

        const item =
            state.inventory.find(
                inventoryItem =>
                    inventoryItem.rowNumber ===
                    rowNumber
            );


        if (!item) {
            return;
        }


        if (
            item.balance !== 0
        ) {

            setInventoryManageStatus(
                `Cannot remove ${item.asset} because its current balance is ${item.balance}. Set the balance to zero first.`,
                true
            );

            return;

        }


        const confirmed =
            window.confirm(
                `Remove the asset type "${item.asset}" from inventory?`
            );


        if (!confirmed) {
            return;
        }


        try {

            setInventoryManageStatus(
                "Removing asset type..."
            );


            const metadata =
                await sheetsGet(
                    `/${encodeURIComponent(
                        state.workbookId
                    )}`
                );


            const sheet =
                (metadata.sheets || [])
                    .find(
                        sheet =>
                            sheet.properties.title ===
                            CONFIG.INVENTORY_SHEET_NAME
                    );


            if (!sheet) {

                throw new Error(
                    `Sheet "${CONFIG.INVENTORY_SHEET_NAME}" was not found.`
                );

            }


            const sheetId =
                sheet.properties.sheetId;


            await sheetsPost(

                `/${encodeURIComponent(
                    state.workbookId
                )}:batchUpdate`,

                {

                    requests: [

                        {

                            deleteDimension: {

                                range: {

                                    sheetId,

                                    dimension:
                                        "ROWS",

                                    startIndex:
                                        rowNumber - 1,

                                    endIndex:
                                        rowNumber

                                }

                            }

                        }

                    ]

                }

            );


            setInventoryManageStatus(
                "Asset type removed successfully."
            );


            await loadDashboard();


        } catch (error) {

            console.error(error);


            setInventoryManageStatus(
                error.message,
                true
            );

        }

    }


    async function renameAssetType(
        rowNumber
    ) {

        const item =
            state.inventory.find(
                inventoryItem =>
                    inventoryItem.rowNumber ===
                    rowNumber
            );


        if (!item) {
            return;
        }


        const newName =
            window.prompt(
                `Rename "${item.asset}" to:`,
                item.asset
            );


        if (
            newName === null
        ) {
            return;
        }


        const cleanName =
            newName.trim();


        if (!cleanName) {

            setInventoryManageStatus(
                "Asset name cannot be empty.",
                true
            );

            return;

        }


        const duplicate =
            state.inventory.some(
                other =>
                    other.rowNumber !==
                        rowNumber &&

                    other.asset
                        .toLowerCase() ===
                    cleanName.toLowerCase()
            );


        if (duplicate) {

            setInventoryManageStatus(
                "Another asset already has that name.",
                true
            );

            return;

        }


        try {

            setInventoryManageStatus(
                "Renaming asset type..."
            );


            const assetCell =
                columnLetter(
                    item.assetColumn
                ) +
                item.rowNumber;


            const range =
                `${quoteSheetName(
                    CONFIG.INVENTORY_SHEET_NAME
                )}!${assetCell}`;


            await sheetsPut(

                `/${encodeURIComponent(
                    state.workbookId
                )}/values/${encodeURIComponent(
                    range
                )}?valueInputOption=USER_ENTERED`,

                {

                    range,

                    majorDimension:
                        "ROWS",

                    values: [[
                        cleanName
                    ]]

                }

            );


            setInventoryManageStatus(
                "Asset type renamed successfully."
            );


            await loadDashboard();


        } catch (error) {

            console.error(error);


            setInventoryManageStatus(
                error.message,
                true
            );

        }

    }


    function openInventoryManager() {

        renderInventoryManager();


        $("inventory-modal")
            .classList
            .remove("hidden");

    }


    function closeInventoryManager() {

        $("inventory-modal")
            .classList
            .add("hidden");

    }


    function renderInventoryManager() {

        const body =
            $("manage-inventory-body");


        if (!body) {
            return;
        }


        body.innerHTML =

            state.inventory.length

                ? state.inventory.map(
                    item => `

                        <tr>

                            <td>

                                <strong>
                                    ${escapeHtml(
                                        item.asset
                                    )}
                                </strong>

                            </td>

                            <td class="num">

                                ${formatNumber(
                                    item.balance
                                )}

                            </td>

                            <td
                                class="actions-cell"
                            >

                                <button
                                    class="table-action"
                                    data-action="rename"
                                    data-row="${item.rowNumber}"
                                >
                                    Rename
                                </button>


                                <button
                                    class="table-action danger"
                                    data-action="delete"
                                    data-row="${item.rowNumber}"
                                >
                                    Remove
                                </button>

                            </td>

                        </tr>

                    `
                ).join("")

                : emptyRow(
                    3,
                    "No asset types configured."
                );

    }


    function handleMovementChange() {

        const movement =
            $("movement").value;


        const client =
            $("client");


        if (
            movement ===
            "DISCARD"
        ) {

            client.value =
                "HSC London (Self)";


            client.disabled =
                true;

        } else {

            client.disabled =
                false;


            if (
                client.value ===
                "HSC London (Self)"
            ) {

                client.value =
                    "";

            }

        }

    }


    function setDefaultTimestamp() {

        const input =
            $("timestamp");


        if (!input) {
            return;
        }


        const now =
            new Date();


        input.value =

            `${String(
                now.getHours()
            ).padStart(2, "0")}:${String(
                now.getMinutes()
            ).padStart(2, "0")}`;

    }


    function signOut() {

        if (
            state.idTokenPayload?.sub
        ) {

            try {

                google.accounts.id.revoke(
                    state.idTokenPayload.sub,
                    () => {}
                );

            } catch (_) {}

        }


        state.idTokenPayload =
            null;

        state.accessToken =
            null;

        state.workbookId =
            null;

        state.workbookName =
            null;

        state.inventory =
            [];

        state.transactions =
            [];

        state.todayLoads =
            [];

        state.clients =
            [];


        $("dashboard")
            .classList
            .add("hidden");


        $("login-card")
            .classList
            .remove("hidden");


        $("google-signin-button")
            .classList
            .remove("hidden");


        $("grant-access")
            .classList
            .add("hidden");


        $("sign-out")
            .classList
            .add("hidden");


        $("user-photo")
            .classList
            .add("hidden");


        $("user-name")
            .textContent =
            "Not signed in";


        $("user-email")
            .textContent =
            "";


        closeInventoryManager();

    }


    function authHeaders() {

        return {

            Authorization:
                `Bearer ${state.accessToken}`

        };

    }


    async function sheetsGet(
        path
    ) {

        return fetchJson(

            SHEETS_API + path,

            {

                headers:
                    authHeaders()

            }

        );

    }


    async function sheetsPost(
        path,
        body
    ) {

        return fetchJson(

            SHEETS_API + path,

            {

                method:
                    "POST",

                headers: {

                    ...authHeaders(),

                    "Content-Type":
                        "application/json"

                },

                body:
                    JSON.stringify(
                        body
                    )

            }

        );

    }


    async function sheetsPut(
        path,
        body
    ) {

        return fetchJson(

            SHEETS_API + path,

            {

                method:
                    "PUT",

                headers: {

                    ...authHeaders(),

                    "Content-Type":
                        "application/json"

                },

                body:
                    JSON.stringify(
                        body
                    )

            }

        );

    }


    async function fetchJson(
        url,
        options
    ) {

        const response =
            await fetch(
                url,
                options
            );


        const text =
            await response.text();


        let data = {};


        try {

            data =
                text
                    ? JSON.parse(text)
                    : {};

        } catch (_) {}


        if (!response.ok) {

            throw new Error(

                data?.error?.message ||
                `Request failed (${response.status})`

            );

        }


        return data;

    }


    function decodeJwtPayload(
        jwt
    ) {

        const part =
            jwt.split(".")[1];


        const normalized =
            part
                .replace(
                    /-/g,
                    "+"
                )
                .replace(
                    /_/g,
                    "/"
                );


        const padded =
            normalized +
            "=".repeat(

                (
                    4 -
                    normalized.length % 4
                ) % 4

            );


        return JSON.parse(

            decodeURIComponent(

                Array.from(
                    atob(padded)
                )

                .map(
                    character =>
                        `%${character
                            .charCodeAt(0)
                            .toString(16)
                            .padStart(2, "0")}`
                )

                .join("")

            )

        );

    }


    function normalizeHeader(
        value
    ) {

        return String(
            value ?? ""
        )
            .trim()
            .toLowerCase()
            .replace(
                /\s+/g,
                " "
            );

    }


    function findColumn(
        header,
        names
    ) {

        const normalized =
            names.map(
                normalizeHeader
            );


        return header.findIndex(
            value =>
                normalized.includes(
                    value
                )
        );

    }


    function numericValue(
        value
    ) {

        const text =
            String(
                value ?? ""
            )
                .replace(
                    /,/g,
                    ""
                )
                .trim();


        if (!text) {
            return 0;
        }


        const number =
            Number(text);


        return Number.isFinite(
            number
        )
            ? number
            : 0;

    }


    function formatNumber(
        value
    ) {

        return Number(
            value || 0
        )
            .toLocaleString(
                "en-GB"
            );

    }


    function quoteSheetName(
        sheetName
    ) {

        return "'" +

            String(
                sheetName
            )
                .replace(
                    /'/g,
                    "''"
                ) +

            "'";

    }


    function columnLetter(
        number
    ) {

        let result =
            "";


        while (
            number > 0
        ) {

            const remainder =
                (
                    number - 1
                ) % 26;


            result =
                String.fromCharCode(
                    65 + remainder
                ) +
                result;


            number =
                Math.floor(
                    (
                        number - 1
                    ) / 26
                );

        }


        return result;

    }


    function escapeDriveQuery(
        value
    ) {

        return String(
            value
        )
            .replace(
                /\\/g,
                "\\\\"
            )
            .replace(
                /'/g,
                "\\'"
            );

    }


    function escapeHtml(
        value
    ) {

        return String(
            value ?? ""
        )
            .replace(
                /[&<>'"]/g,
                character =>
                    ({

                        "&":
                            "&amp;",

                        "<":
                            "&lt;",

                        ">":
                            "&gt;",

                        "'":
                            "&#39;",

                        '"':
                            "&quot;"

                    }[
                        character
                    ])

            );

    }


    function escapeAttr(
        value
    ) {

        return escapeHtml(
            value
        );

    }


    function emptyRow(
        span,
        text
    ) {

        return `

            <tr>

                <td
                    colspan="${span}"
                    class="empty"
                >
                    ${escapeHtml(
                        text
                    )}
                </td>

            </tr>

        `;

    }


    function setAuthStatus(
        text,
        error = false
    ) {

        const element =
            $("auth-status");


        element.textContent =
            text;


        element.className =
            `status ${
                error
                    ? "error"
                    : ""
            }`;

    }


    function setMovementStatus(
        text,
        error = false
    ) {

        const element =
            $("movement-status");


        element.textContent =
            text;


        element.className =
            `status ${
                error
                    ? "error"
                    : ""
            }`;

    }


    function setInventoryManageStatus(
        text,
        error = false
    ) {

        const element =
            $("inventory-manage-status");


        element.textContent =
            text;


        element.className =
            `status ${
                error
                    ? "error"
                    : ""
            }`;

    }


    function setSyncStatus(
        text,
        error = false
    ) {

        const element =
            $("sync-status");


        element.textContent =
            text;


        element.className =
            `muted ${
                error
                    ? "error"
                    : ""
            }`;

    }


})();
