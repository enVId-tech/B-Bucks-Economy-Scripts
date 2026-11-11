/* Coded by Erick Tran for Mr. Banderas, 2026
 * Copyright (c) 2026 Erick Tran.
 * This file is licensed under the MIT License, check the LICENSE file for details.
 *
 * GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
 *
 * !!!--- IMPORTANT NOTE: This file was originally written in TypeScript, go to the GitHub to see the original non-compiled code. ---!!!
 *
 * Server side of the student portal, a web app that lets every student open ONE link, sign in with their school Google account and
 * see their own balance, invest, request expenditures and watch the class leaderboard. Admins (people who can edit the spreadsheet)
 * open the same link and can go into any student's account to act for them and approve or deny expenditure requests.
 *
 * Security model, read before changing anything:
 *  - The web app is deployed as "Execute as: me" so students never need access to the spreadsheet. Who is calling is therefore taken
 *    ONLY from Session.getActiveUser(), never from anything the browser sends.
 *  - Every endpoint re-checks the caller on the server. The browser is only a view.
 *  - A student can only act on the profiles linked to their own email. Admins can open any student's account and act on it exactly like the student would (same rules, no override); each such action is written to the log with the admin's email.
 *  - The allowed email domain comes from a Script Property (getPortalAllowedDomain), it is never written in the source.
 *  - Money only moves through the existing handlers (handleDeposit, handleWithdraw, applyMathToSelection) so every transaction
 *    still lands in the log, takes the transaction lock and respects the same caps as the admin dialogs.
 */

/** Messages that log(..., true) would normally show as a spreadsheet alert. While a portal request runs they are collected here. */
let portalCapturedMessages: string[] | null = null;

type PortalStatus = "ok" | "signin" | "wrongDomain" | "notConfigured";

interface PortalIdentity {
    status: PortalStatus;
    email: string;
    isAdmin: boolean;
}

/** One student account the signed-in person may act as: a student name on one period sheet. */
interface PortalProfile {
    name: string;
    period: string;
}

interface PortalStudentRow {
    name: string;
    /** From the hidden ID column. Used only to match a sign-in email, never sent to the browser. */
    studentId: string;
    row: number;
    balance: number;
    expenditures: number;
    initialDeposit: number;
    depositDate: Date | null;
    grossGain: number;
    netGain: number;
    netPercentageGain: number;
}

interface PortalActionResult {
    ok: boolean;
    message: string;
}

interface PortalRequestRecord {
    id: number;
    requestedAt: string;
    email: string;
    student: string;
    period: string;
    itemName: string;
    quarter: string;
    unitPrice: number;
    quantity: number;
    total: number;
    status: "Pending" | "Approved" | "Denied" | "Cancelled";
    resolvedBy: string;
    resolvedAt: string;
    note: string;
}

const PORTAL_REQUEST_HEADERS: string[] = ["ID", "Requested At", "Email", "Student", "Period", "Item", "Quarter", "Unit Price", "Quantity", "Total", "Status", "Resolved By", "Resolved At", "Note"];
const PORTAL_ACCOUNT_HEADERS: string[] = ["Email", "Period", "Name"];
const PORTAL_ADMIN_CACHE_KEY: string = "portalAdminEmails";
const PORTAL_MAX_REQUEST_QUANTITY: number = 50;

// ---------------------------------------------------------------------------------------------------------------------------
// Web app entry point and setup
// ---------------------------------------------------------------------------------------------------------------------------

/** Web app entry point: serves the portal page. The page asks portalBootstrap() who the visitor is once it has loaded. */
function doGet(): GoogleAppsScript.HTML.HtmlOutput {
    const template = HtmlService.createTemplateFromFile("Portal");
    return template.evaluate()
        .setTitle("Bank of Banderas")
        .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/** Menu item: asks for the email domain students must sign in with and stores it as a Script Property. */
function configurePortalDomain(): void {
    const ui = SpreadsheetApp.getUi();
    const currentDomain = getPortalAllowedDomain();
    const answer = ui.prompt(
        "Student portal domain",
        `Enter the email domain students sign in with (the part after the @).${currentDomain ? `\nCurrent value: ${currentDomain}` : "\nNothing is set yet, so only spreadsheet editors can open the portal."}\nLeave empty to clear it.`,
        ui.ButtonSet.OK_CANCEL
    );
    if (answer.getSelectedButton() !== ui.Button.OK) return;

    const enteredDomain = answer.getResponseText().trim().toLowerCase().replace(/^@/, "");
    if (enteredDomain && !/^[a-z0-9.-]+\.[a-z]{2,}$/.test(enteredDomain)) {
        ui.alert(`"${enteredDomain}" is not a valid domain. Nothing was changed.`);
        return;
    }

    const properties = PropertiesService.getScriptProperties();
    if (enteredDomain) properties.setProperty(PORTAL_DOMAIN_PROPERTY, enteredDomain);
    else properties.deleteProperty(PORTAL_DOMAIN_PROPERTY);
    ui.alert(enteredDomain ? "The student portal now accepts that domain." : "The portal domain was cleared.");
}

/** Menu item: creates the "Portal Accounts" and "Expenditure Requests" tabs (with headers) when they do not exist yet. */
function createPortalSheets(): void {
    ensurePortalSheet_(PORTAL_ACCOUNTS_SHEET, PORTAL_ACCOUNT_HEADERS);
    ensurePortalSheet_(PORTAL_REQUESTS_SHEET, PORTAL_REQUEST_HEADERS);
    SpreadsheetApp.getActive().toast("Portal tabs are ready.", "Student portal", 5);
}

/** Returns the named portal tab, creating it with a bold header row when needed. */
function ensurePortalSheet_(sheetName: string, headers: string[]): GoogleAppsScript.Spreadsheet.Sheet {
    const existingSheet = getSheetNamed(sheetName);
    if (existingSheet) return existingSheet;

    const sheet = getSS().insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    sheet.setFrozenRows(1);
    resetSheetMemo();
    return sheet;
}

// ---------------------------------------------------------------------------------------------------------------------------
// Who is calling
// ---------------------------------------------------------------------------------------------------------------------------

/** Emails of everyone who can edit the spreadsheet (owner included). They are the portal admins. */
function getPortalAdminEmails_(): string[] {
    const cache = CacheService.getScriptCache();
    const cachedEmails = cache.get(PORTAL_ADMIN_CACHE_KEY);
    if (cachedEmails) return JSON.parse(cachedEmails);

    const spreadsheet = getSS();
    const adminEmails = new Set<string>();
    spreadsheet.getEditors().forEach(editor => adminEmails.add(editor.getEmail().toLowerCase()));
    const owner = spreadsheet.getOwner();
    if (owner) adminEmails.add(owner.getEmail().toLowerCase());
    const effectiveUser = Session.getEffectiveUser().getEmail();
    if (effectiveUser) adminEmails.add(effectiveUser.toLowerCase());

    const emailList = Array.from(adminEmails).filter(email => email);
    cache.put(PORTAL_ADMIN_CACHE_KEY, JSON.stringify(emailList), 600);
    return emailList;
}

/** Works out who is calling from Google's own session, never from browser input. */
function resolvePortalIdentity_(): PortalIdentity {
    const email = (Session.getActiveUser().getEmail() || "").trim().toLowerCase();
    if (!email) return { status: "signin", email: "", isAdmin: false };

    const isAdmin = getPortalAdminEmails_().includes(email);
    const allowedDomain = getPortalAllowedDomain();
    if (!allowedDomain) return { status: isAdmin ? "ok" : "notConfigured", email, isAdmin };
    if (!isAdmin && !email.endsWith(`@${allowedDomain}`)) return { status: "wrongDomain", email, isAdmin };
    return { status: "ok", email, isAdmin };
}

/** Link that sends a visitor without a Google session to the Google account chooser and back to the portal. */
function buildPortalSignInUrl_(): string {
    const portalUrl = ScriptApp.getService().getUrl();
    return `https://accounts.google.com/AccountChooser?continue=${encodeURIComponent(portalUrl)}`;
}

/** Turns "Last, First M" into "First L." for the leaderboard. Anything without a comma is shown as it is. */
function abbreviateStudentName_(fullName: string): string {
    const commaIndex = fullName.indexOf(",");
    if (commaIndex < 0) return fullName;
    const lastName = fullName.slice(0, commaIndex).trim();
    const firstName = fullName.slice(commaIndex + 1).trim().split(/\s+/)[0] ?? "";
    return `${firstName} ${lastName.charAt(0)}.`.trim();
}

// ---------------------------------------------------------------------------------------------------------------------------
// Rosters and profile linking
// ---------------------------------------------------------------------------------------------------------------------------

/** Every period tab, the same way the investment ledger finds them. */
function getPortalPeriodSheets_(): GoogleAppsScript.Spreadsheet.Sheet[] {
    return getAllSheets().filter(sheet => sheet.getName().includes("Period"));
}

function toNumber_(rawValue: any): number {
    const parsedNumber = Number(rawValue);
    return Number.isFinite(parsedNumber) ? parsedNumber : 0;
}

/** Reads every student of one period tab with a single range call. */
function readPortalPeriod_(sheet: GoogleAppsScript.Spreadsheet.Sheet): PortalStudentRow[] {
    const values = sheet.getDataRange().getValues();
    const students: PortalStudentRow[] = [];

    for (let rowIndex = getUserStartingRow() - 1; rowIndex < values.length; rowIndex++) {
        const row = values[rowIndex];
        const name = String(row[getNamesColumn() - 1] ?? "").trim();
        if (!name) continue;

        const depositValue = row[getDepositDateColumn() - 1];
        const depositDate = depositValue instanceof Date && !isNaN(depositValue.getTime()) ? depositValue : null;

        students.push({
            name,
            studentId: String(row[getStudentIdColumn() - 1] ?? "").trim(),
            row: rowIndex + 1,
            balance: Number(toNumber_(row[getBalanceColumn() - 1]).toFixed(2)),
            expenditures: toNumber_(row[getExpendituresColumn() - 1]),
            initialDeposit: Number(toNumber_(row[getInitialDepositColumn() - 1]).toFixed(2)),
            depositDate,
            grossGain: Number(toNumber_(row[getGrossInvestmentGainColumn() - 1]).toFixed(2)),
            netGain: Number(toNumber_(row[getNetInvestmentGainColumn() - 1]).toFixed(2)),
            netPercentageGain: Number(toNumber_(row[getNetPercentageGainColumn() - 1]).toFixed(4))
        });
    }
    return students;
}

/** "3" and "Period 3" both mean the tab named "Period 3". */
function normalizePortalPeriodName_(rawPeriod: any): string {
    const text = String(rawPeriod ?? "").trim();
    return /^\d+$/.test(text) ? `Period ${text}` : text;
}

/** Profiles listed for this email in the "Portal Accounts" tab (an optional manual override, for example for an email that does not contain the ID). A row only counts when the student really is on that period tab. */
function findListedProfiles_(email: string): PortalProfile[] {
    const accountsSheet = getSheetNamed(PORTAL_ACCOUNTS_SHEET);
    if (!accountsSheet || accountsSheet.getLastRow() < 2) return [];

    const profiles: PortalProfile[] = [];
    const rosterByPeriod = new Map<string, Set<string>>();
    const rows = accountsSheet.getRange(2, 1, accountsSheet.getLastRow() - 1, PORTAL_ACCOUNT_HEADERS.length).getValues();

    rows.forEach(row => {
        if (String(row[0] ?? "").trim().toLowerCase() !== email) return;
        const period = normalizePortalPeriodName_(row[1]);
        const name = String(row[2] ?? "").trim();
        const periodSheet = getSheetNamed(period);
        if (!periodSheet || !name) return;

        if (!rosterByPeriod.has(period)) {
            rosterByPeriod.set(period, new Set(readPortalPeriod_(periodSheet).map(student => student.name.toLowerCase())));
        }
        if (rosterByPeriod.get(period)!.has(name.toLowerCase())) profiles.push({ name, period });
    });
    return profiles;
}

/**
 * Matches an email to students by student ID. School addresses contain the student's ID (for example 123456@school.example or jdoe123456@...),
 * so the email's local part is split into runs of digits and a student matches when one run is exactly their ID.
 * A student who is on several period tabs has the same ID on each, so they get one profile per tab.
 */
function findIdMatchedProfiles_(email: string): PortalProfile[] {
    const digitRuns = new Set(email.split("@")[0].match(/\d+/g) ?? []);
    if (digitRuns.size === 0) return [];

    const matches: PortalProfile[] = [];
    getPortalPeriodSheets_().forEach(sheet => {
        readPortalPeriod_(sheet).forEach(student => {
            if (student.studentId && digitRuns.has(student.studentId)) matches.push({ name: student.name, period: sheet.getName() });
        });
    });
    return matches;
}

/** Every student account the signed-in person may act as. */
function findPortalProfiles_(identity: PortalIdentity): PortalProfile[] {
    const listedProfiles = findListedProfiles_(identity.email);
    return listedProfiles.length > 0 ? listedProfiles : findIdMatchedProfiles_(identity.email);
}

/** True when the caller may use this profile. */
function ownsPortalProfile_(profiles: PortalProfile[], name: string, period: string): boolean {
    return profiles.some(profile => profile.name === name && profile.period === period);
}

/** Runs an action while collecting the messages log(..., true) would have shown as spreadsheet alerts. */
function runCapturingMessages_<T>(work: () => T): { result: T; messages: string[] } {
    const previousMessages = portalCapturedMessages;
    const collectedMessages: string[] = [];
    portalCapturedMessages = collectedMessages;
    try {
        return { result: work(), messages: collectedMessages };
    } finally {
        portalCapturedMessages = previousMessages;
    }
}

function failure_(message: string): PortalActionResult {
    return { ok: false, message };
}

// ---------------------------------------------------------------------------------------------------------------------------
// Reading data for the page
// ---------------------------------------------------------------------------------------------------------------------------

/**
 * First call of the page: who is this, and which student accounts can they open?
 * Admins also get the list of every student so they can spectate.
 */
function portalBootstrap(): any {
    try {
        const identity = resolvePortalIdentity_();
        const response: any = { status: identity.status, email: identity.email, isAdmin: identity.isAdmin, signInUrl: buildPortalSignInUrl_(), profiles: [] };
        if (identity.status !== "ok") return response;

        response.profiles = findPortalProfiles_(identity);
        if (identity.isAdmin) response.directory = buildPortalDirectory_();
        return response;
    } catch (error: any) {
        log(`Error in portalBootstrap: ${error.message}`, false);
        return { status: "error", message: "The portal could not load. Please try again in a moment." };
    }
}

/** Period tabs with their student names, used by the admin student picker. */
function buildPortalDirectory_(): { period: string; students: string[] }[] {
    return getPortalPeriodSheets_().map(sheet => ({ period: sheet.getName(), students: readPortalPeriod_(sheet).map(student => student.name) }));
}

interface PortalPurchasableItem {
    itemName: string;
    unitPrice: number;
    personRemaining: number | null;
    personLimit: number | null;
    classRemaining: number | null;
    classLimit: number | null;
}

/** Units of an item that are already waiting in pending requests, by student and by whole class. */
function countPendingUnits_(itemName: string, quarter: string, period: string, student: string, ignoreRequestId?: number): { person: number; classWide: number } {
    let person = 0;
    let classWide = 0;
    readPortalRequests_().forEach(request => {
        if (request.status !== "Pending" || request.itemName !== itemName || request.quarter !== quarter || request.period !== period) return;
        if (request.id === ignoreRequestId) return;
        classWide += request.quantity;
        if (request.student === student) person += request.quantity;
    });
    return { person, classWide };
}

/** What the student may buy this quarter, with the caps that are left (used purchases and pending requests both count). */
function buildPortalCatalog_(quarter: QuarterId, studentName: string, period: string): PortalPurchasableItem[] {
    const services = fetchServicesDataCached();
    if (!Array.isArray(services)) return [];

    const enforced = isServicesLimitEnforced();
    const periodNumber = parseInt(period.replace(/\D/g, ""), 10);
    const items: PortalPurchasableItem[] = [];

    services.forEach(service => {
        const unitPrice = service.pricing?.[quarter];
        if (service.category !== "Expense" || typeof unitPrice !== "number" || !(unitPrice > 0)) return;

        const { personLimit, classLimit } = lookupItemLimits(service.itemName, quarter);
        let personRemaining: number | null = null;
        let classRemaining: number | null = null;

        if (enforced && (personLimit !== null || classLimit !== null)) {
            const counts = countPurchasesFromLog(service.itemName, quarter);
            const pendingUnits = countPendingUnits_(service.itemName, quarter, period, studentName);

            if (personLimit !== null) {
                const used = counts.get(`${studentName}|${Number.isFinite(periodNumber) ? periodNumber : ""}`) ?? 0;
                personRemaining = Math.max(0, personLimit - used - pendingUnits.person);
            }
            if (classLimit !== null) {
                let classUsed = 0;
                counts.forEach((units, key) => {
                    if (Number(key.slice(key.lastIndexOf("|") + 1)) === periodNumber) classUsed += units;
                });
                classRemaining = Math.max(0, classLimit - classUsed - pendingUnits.classWide);
            }
        }

        items.push({ itemName: service.itemName, unitPrice, personLimit: personLimit ?? null, personRemaining, classLimit: classLimit ?? null, classRemaining });
    });
    return items;
}

/** Checks a purchase against price, balance and caps. Returns the reason it is not allowed, or null. */
function checkPortalPurchase_(student: PortalStudentRow, period: string, itemName: string, quarter: QuarterId, quantity: number, ignoreRequestId?: number): string | null {
    const item = buildPortalCatalogForCheck_(quarter, student.name, period, ignoreRequestId).find(candidate => candidate.itemName === itemName);
    if (!item) return `"${itemName}" is not for sale in ${quarter}.`;

    const total = Number((item.unitPrice * quantity).toFixed(2));
    if (student.balance - total < 0) return `That costs $${total.toFixed(2)} but your balance is $${student.balance.toFixed(2)}.`;
    if (item.personRemaining !== null && quantity > item.personRemaining) return `Only ${item.personRemaining} left for you this quarter (max per person reached).`;
    if (item.classRemaining !== null && quantity > item.classRemaining) return `Only ${item.classRemaining} left for your class this quarter (max per class reached).`;
    return null;
}

/** buildPortalCatalog_ with one request left out of the pending count, so approving a request does not count it against itself. */
function buildPortalCatalogForCheck_(quarter: QuarterId, studentName: string, period: string, ignoreRequestId?: number): PortalPurchasableItem[] {
    const catalog = buildPortalCatalog_(quarter, studentName, period);
    if (ignoreRequestId === undefined) return catalog;

    const ownRequest = readPortalRequests_().find(request => request.id === ignoreRequestId && request.status === "Pending");
    if (!ownRequest) return catalog;
    return catalog.map(item => {
        if (item.itemName !== ownRequest.itemName) return item;
        return {
            ...item,
            personRemaining: item.personRemaining === null ? null : item.personRemaining + (ownRequest.student === studentName ? ownRequest.quantity : 0),
            classRemaining: item.classRemaining === null ? null : item.classRemaining + ownRequest.quantity
        };
    });
}

/** The quarter the portal prices items in (the same one the Services dialog selects). */
function getPortalQuarter_(): QuarterId {
    const quarterInfo = getCurrentQuarterInfo();
    return 'error' in quarterInfo ? "Q1" : quarterInfo.quarter;
}

function weeksUntilInvestingEnds_(): number {
    const endInvestingDate = fetchProperty("endInvestingDate", "importantDates");
    if (!endInvestingDate || typeof endInvestingDate === "object") return 0;
    const endMs = new Date(endInvestingDate).getTime();
    if (isNaN(endMs)) return 0;
    return Math.max(0, Math.floor((endMs - Date.now()) / (1000 * 60 * 60 * 24 * 7)));
}

/** Days still to wait before the CoD can be withdrawn (0 when it can be taken out, or when there is no CoD). */
function daysLeftInLockout_(student: PortalStudentRow): number {
    const lockoutSetting = Number(fetchProperty("withdrawLockoutTimeframeDays", "mandatedPolicies"));
    const lockoutDays = Number.isFinite(lockoutSetting) && lockoutSetting >= 0 ? lockoutSetting : 14;
    if (!student.depositDate) return 0;
    const elapsedDays = (Date.now() - student.depositDate.getTime()) / (1000 * 60 * 60 * 24);
    return Math.max(0, Math.ceil(lockoutDays - elapsedDays));
}

/** Is it currently allowed to invest? Open when no investing dates are set. */
function investingWindowMessage_(): string | null {
    const startValue = fetchProperty("startInvestingDate", "importantDates");
    const endValue = fetchProperty("endInvestingDate", "importantDates");
    const startMs = startValue && typeof startValue !== "object" ? new Date(startValue).getTime() : NaN;
    const endMs = endValue && typeof endValue !== "object" ? new Date(endValue).getTime() : NaN;
    const now = Date.now();
    if (!isNaN(startMs) && now < startMs) return "Investing has not opened yet.";
    if (!isNaN(endMs) && now > endMs + 24 * 60 * 60 * 1000) return "Investing is closed for this period.";
    return null;
}

/**
 * Everything the dashboard shows for one student.
 * @param data JSON string: { name: string, period: string }. Students may only ask for their own profile, admins for any.
 */
function portalGetDashboard(data: string): any {
    try {
        const request = JSON.parse(data);
        const name = String(request.name ?? "");
        const period = String(request.period ?? "");

        const identity = resolvePortalIdentity_();
        if (identity.status !== "ok") return { error: "You are not signed in." };

        const profiles = findPortalProfiles_(identity);
        const isOwnProfile = ownsPortalProfile_(profiles, name, period);
        if (!isOwnProfile && !identity.isAdmin) return { error: "That account is not linked to your email." };

        const periodSheet = getSheetNamed(period);
        if (!periodSheet) return { error: `Period "${period}" was not found.` };
        const classmates = readPortalPeriod_(periodSheet);
        const student = classmates.find(candidate => candidate.name === name);
        if (!student) return { error: `${name} was not found on ${period}.` };

        const quarter = getPortalQuarter_();
        const lockoutDaysLeft = student.netGain !== 0 || student.initialDeposit > 0 ? daysLeftInLockout_(student) : 0;
        const hasActiveCod = student.initialDeposit > 0 || student.netGain !== 0;
        const investingMessage = investingWindowMessage_();

        const leaderboard = classmates
            .map(classmate => ({ name: classmate.name, total: Number((classmate.balance + (classmate.netGain > 0 ? classmate.netGain : 0)).toFixed(2)) }))
            .sort((first, second) => second.total - first.total)
            .map((entry, index) => ({ rank: index + 1, displayName: abbreviateStudentName_(entry.name), total: entry.total, isYou: entry.name === name }));

        const weeklyInterestRate = Number(fetchProperty("weeklyInterestRate", "standardPercentages"));

        return {
            viewingAsAdmin: !isOwnProfile,
            student: {
                name, period,
                displayName: abbreviateStudentName_(name),
                balance: student.balance,
                expenditures: student.expenditures,
                codPrincipal: student.initialDeposit,
                codDate: student.depositDate ? student.depositDate.toISOString() : null,
                codGross: student.grossGain,
                codNet: student.netGain,
                codPercentageGain: student.netPercentageGain,
                hasActiveCod,
                lockoutDaysLeft,
                canWithdraw: hasActiveCod && student.netGain !== 0 && lockoutDaysLeft === 0,
                canDeposit: !investingMessage && !(hasActiveCod && lockoutDaysLeft > 0),
                depositBlockedReason: investingMessage ?? (hasActiveCod && lockoutDaysLeft > 0 ? `Your CoD is locked for ${lockoutDaysLeft} more day(s). You can invest again once it unlocks.` : ""),
                availableToInvest: Number((student.balance + (student.netGain > 0 && lockoutDaysLeft === 0 ? student.netGain : 0)).toFixed(2))
            },
            predictor: {
                weeklyInterestRate: Number.isFinite(weeklyInterestRate) ? weeklyInterestRate : null,
                weeksRemaining: weeksUntilInvestingEnds_(),
                currentCodAmount: Number((student.balance + student.netGain).toFixed(2))
            },
            quarter,
            catalog: buildPortalCatalog_(quarter, name, period),
            requests: readPortalRequests_().filter(candidate => candidate.student === name && candidate.period === period).slice(-20).reverse(),
            purchases: fetchPortalUndoablePurchases_(name, period),
            leaderboard
        };
    } catch (error: any) {
        log(`Error in portalGetDashboard: ${error.message}`, false);
        return { error: "Your dashboard could not load. Please try again in a moment." };
    }
}

/** The log filters the portal offers, mapped to the log's own type labels. */
const PORTAL_ACTIVITY_CATEGORIES: { [category: string]: string } = { income: "income", spending: "spending", investments: "deposit" };

/**
 * One page of a student's transaction history, newest first, read from the transactions sheet named in Settings. Every transaction
 * of the student is included. The text search looks at the description (what was bought, invested or earned).
 * @param data JSON string: { name, period, page?, search?, category?: "all" | "income" | "spending" | "investments" }
 */
function portalGetTransactions(data: string): any {
    try {
        const request = JSON.parse(data);
        const name = String(request.name ?? "");
        const period = String(request.period ?? "");
        const search = String(request.search ?? "").trim().slice(0, 80);
        const category = String(request.category ?? "all");
        const requestedPage = Math.max(1, Math.floor(Number(request.page)) || 1);

        const identity = resolvePortalIdentity_();
        if (identity.status !== "ok") return { error: "You are not signed in." };
        if (!ownsPortalProfile_(findPortalProfiles_(identity), name, period) && !identity.isAdmin) return { error: "That account is not linked to your email." };

        const filters: TransactionFilters = { student: name, period: period.replace(/\D/g, "") };
        if (search) { filters.search = search; filters.scope = "item"; }
        if (PORTAL_ACTIVITY_CATEGORIES[category]) filters.type = PORTAL_ACTIVITY_CATEGORIES[category];

        const response = queryTransactions(JSON.stringify({ page: requestedPage, pageSize: PORTAL_RECENT_TRANSACTIONS, filters }));
        if ('error' in response) {
            log(`Portal: the transaction history could not be read: ${response.error}`, false);
            return { error: `The transaction history is not available (the "${getTransactionsSheetName()}" tab could not be read).` };
        }

        const firstOnPage = (response.page - 1) * response.pageSize - response.window.startIndex;
        const records = response.window.records
            .slice(Math.max(0, firstOnPage), Math.max(0, firstOnPage) + response.pageSize)
            .map(record => ({
                id: record.id,
                timestamp: record.timestamp,
                type: record.type,
                description: (record.serviceProvided ?? "").trim(),
                amount: Math.abs(Number(record.tenderedMoney) || 0),
                change: transactionBalanceChange(record)
            }));

        return { records, total: response.total, page: response.page, totalPages: response.totalPages };
    } catch (error: any) {
        log(`Error in portalGetTransactions: ${error.message}`, false);
        return { error: "The transaction history could not load. Please try again." };
    }
}

/**
 * The student's purchases that can still be undone: purchase records that are still in effect (not undone) and still on the log.
 * A purchase stops being listed once its row is deleted from the transaction log.
 */
function fetchPortalUndoablePurchases_(name: string, period: string): any[] {
    const response = queryTransactions(JSON.stringify({
        page: 1,
        pageSize: PORTAL_PURCHASE_SCAN_SIZE,
        filters: { student: name, period: period.replace(/\D/g, "") }
    }));
    if ('error' in response) return [];

    const ownRecords = response.window.records.filter(record => record.individual === name);
    // The window is newest first, the in-effect walk needs oldest first.
    const inEffect = resolveRecordsInEffect([...ownRecords].reverse());
    return ownRecords
        .filter(record => parsePurchaseItemName(record.serviceProvided) !== null && inEffect.has(Number(record.id)))
        .map(record => ({
            id: record.id,
            timestamp: record.timestamp,
            description: (record.serviceProvided ?? "").trim(),
            amount: Number(((record.unitPrice ?? 0) * (record.quantity ?? 1)).toFixed(2))
        }));
}

/**
 * Takes back a purchase. Works until the purchase's row is deleted from the transaction log; the money goes back and the
 * quantity no longer counts against the caps. The same undo as the admin Transaction Records dialog runs.
 * @param data JSON string: { name, period, transactionId }
 */
function portalUndoPurchase(data: string): PortalActionResult {
    return runWithTransactionLock(() => {
        try {
            const request = JSON.parse(data);
            const name = String(request.name ?? "");
            const period = String(request.period ?? "");
            const transactionId = Number(request.transactionId);

            const { identity, reason, onBehalf } = authorizeProfileAction_(name, period);
            if (reason) return failure_(reason);
            if (!Number.isInteger(transactionId)) return failure_("That purchase was not found.");

            const record = fetchTransactionsByIds([transactionId])[0];
            if (!record) return failure_("That purchase is no longer on the transaction log, so it cannot be undone.");
            if (record.individual !== name || String(record.period) !== period.replace(/\D/g, "")) return failure_("That purchase was not found.");
            if (parsePurchaseItemName(record.serviceProvided) === null) return failure_("Only purchases can be undone here.");

            const { result, messages } = runCapturingMessages_(() => undoTransactions(JSON.stringify({ transactionIds: [transactionId] })));
            if (result !== true) return failure_(messages.join(" ") || "The purchase could not be undone.");

            if (onBehalf) log(`Portal: ${identity.email} undid transaction ${transactionId} for ${name} (${period}).`, false);
            return { ok: true, message: "Purchase undone. The money is back in the balance." };
        } catch (error: any) {
            log(`Error in portalUndoPurchase: ${error.message}`, false);
            return failure_("The purchase could not be undone. Please try again.");
        }
    }, failure_("The bank is busy right now. Please try again in a moment."));
}

// ---------------------------------------------------------------------------------------------------------------------------
// Student actions: invest, withdraw, request expenditures
// ---------------------------------------------------------------------------------------------------------------------------

/**
 * Confirms the caller is signed in and may ACT as this profile: their own, or any student's when they are an admin.
 * `onBehalf` is true when an admin is acting on somebody else's account. The same rules and caps apply either way.
 */
function authorizeProfileAction_(name: string, period: string): { identity: PortalIdentity; reason: string | null; onBehalf: boolean } {
    const identity = resolvePortalIdentity_();
    if (identity.status !== "ok") return { identity, reason: "You are not signed in.", onBehalf: false };
    if (ownsPortalProfile_(findPortalProfiles_(identity), name, period)) return { identity, reason: null, onBehalf: false };
    if (identity.isAdmin) {
        const exists = findPortalStudent_(name, period) !== null;
        return { identity, reason: exists ? null : "That student was not found.", onBehalf: exists };
    }
    return { identity, reason: "That account is not linked to your email.", onBehalf: false };
}

function findPortalStudent_(name: string, period: string): PortalStudentRow | null {
    const periodSheet = getSheetNamed(period);
    if (!periodSheet) return null;
    return readPortalPeriod_(periodSheet).find(student => student.name === name) ?? null;
}

/**
 * Invests cash into the student's CoD. The same handler as the admin dialog runs, always WITHOUT the admin override.
 * @param data JSON string: { name, period, amount }
 */
function portalDeposit(data: string): PortalActionResult {
    try {
        const request = JSON.parse(data);
        const name = String(request.name ?? "");
        const period = String(request.period ?? "");
        const amount = Number(Number(request.amount).toFixed(2));

        const { identity, reason, onBehalf } = authorizeProfileAction_(name, period);
        if (reason) return failure_(reason);
        if (!Number.isFinite(amount) || amount <= 0) return failure_("Enter an amount greater than zero.");

        const windowMessage = investingWindowMessage_();
        if (windowMessage) return failure_(windowMessage);

        const { result, messages } = runCapturingMessages_(() => handleDeposit(JSON.stringify({ amount, student: name, period, override: false })));
        if (result === true) {
            if (onBehalf) log(`Portal: ${identity.email} invested $${amount.toFixed(2)} for ${name} (${period}).`, false);
            return { ok: true, message: `Invested $${amount.toFixed(2)}.` };
        }
        return failure_(messages.join(" ") || "The deposit could not be completed.");
    } catch (error: any) {
        log(`Error in portalDeposit: ${error.message}`, false);
        return failure_("The deposit could not be completed. Please try again.");
    }
}

/**
 * Takes the CoD out once the withdrawal lockout is over. The lockout is only checked in the admin dialog's page code, so it is enforced here.
 * @param data JSON string: { name, period }
 */
function portalWithdraw(data: string): PortalActionResult {
    try {
        const request = JSON.parse(data);
        const name = String(request.name ?? "");
        const period = String(request.period ?? "");

        const { identity, reason, onBehalf } = authorizeProfileAction_(name, period);
        if (reason) return failure_(reason);

        const student = findPortalStudent_(name, period);
        if (!student) return failure_("Your account was not found.");
        if (student.netGain === 0) return failure_("You have no CoD to withdraw.");

        const daysLeft = daysLeftInLockout_(student);
        if (daysLeft > 0) return failure_(`Your CoD is locked for ${daysLeft} more day(s).`);

        const { result, messages } = runCapturingMessages_(() => handleWithdraw(JSON.stringify({ amount: student.netGain, student: name, period, override: false })));
        if (result === true) {
            if (onBehalf) log(`Portal: ${identity.email} withdrew $${student.netGain.toFixed(2)} for ${name} (${period}).`, false);
            return { ok: true, message: `Withdrew $${student.netGain.toFixed(2)} to the balance.` };
        }
        return failure_(messages.join(" ") || "The withdrawal could not be completed.");
    } catch (error: any) {
        log(`Error in portalWithdraw: ${error.message}`, false);
        return failure_("The withdrawal could not be completed. Please try again.");
    }
}

/** All requests in the "Expenditure Requests" tab. */
function readPortalRequests_(): PortalRequestRecord[] {
    const requestsSheet = getSheetNamed(PORTAL_REQUESTS_SHEET);
    if (!requestsSheet || requestsSheet.getLastRow() < 2) return [];

    return requestsSheet.getRange(2, 1, requestsSheet.getLastRow() - 1, PORTAL_REQUEST_HEADERS.length).getValues()
        .filter(row => row[0] !== "" && row[0] !== null)
        .map(row => ({
            id: Number(row[0]),
            requestedAt: row[1] instanceof Date ? row[1].toISOString() : String(row[1]),
            email: String(row[2]),
            student: String(row[3]),
            period: String(row[4]),
            itemName: String(row[5]),
            quarter: String(row[6]),
            unitPrice: Number(row[7]),
            quantity: Number(row[8]),
            total: Number(row[9]),
            status: String(row[10]) as PortalRequestRecord["status"],
            resolvedBy: String(row[11] ?? ""),
            resolvedAt: row[12] instanceof Date ? row[12].toISOString() : String(row[12] ?? ""),
            note: String(row[13] ?? "")
        }));
}

/** Row number (1-based) of a request in its tab, or -1. */
function findPortalRequestRow_(requestId: number): number {
    const requestsSheet = getSheetNamed(PORTAL_REQUESTS_SHEET);
    if (!requestsSheet || requestsSheet.getLastRow() < 2) return -1;
    const ids = requestsSheet.getRange(2, 1, requestsSheet.getLastRow() - 1, 1).getValues();
    const index = ids.findIndex(row => Number(row[0]) === requestId);
    return index < 0 ? -1 : index + 2;
}

/**
 * A student asks to buy something. Nothing is charged until an admin approves it.
 * @param data JSON string: { name, period, itemName, quantity }
 */
function portalRequestExpenditure(data: string): PortalActionResult {
    return runWithTransactionLock(() => portalRequestExpenditureInsideLock_(data), failure_("The bank is busy right now. Please try again in a moment."));
}

function portalRequestExpenditureInsideLock_(data: string): PortalActionResult {
    try {
        const request = JSON.parse(data);
        const name = String(request.name ?? "");
        const period = String(request.period ?? "");
        const itemName = String(request.itemName ?? "").trim();
        const quantity = Math.floor(Number(request.quantity));

        const { identity, reason, onBehalf } = authorizeProfileAction_(name, period);
        if (reason) return failure_(reason);
        if (!itemName) return failure_("Choose something to request.");
        if (!Number.isFinite(quantity) || quantity < 1 || quantity > PORTAL_MAX_REQUEST_QUANTITY) return failure_(`Quantity must be between 1 and ${PORTAL_MAX_REQUEST_QUANTITY}.`);

        const student = findPortalStudent_(name, period);
        if (!student) return failure_("Your account was not found.");

        const quarter = getPortalQuarter_();
        const refusal = checkPortalPurchase_(student, period, itemName, quarter, quantity);
        if (refusal) return failure_(refusal);

        const item = buildPortalCatalog_(quarter, name, period).find(candidate => candidate.itemName === itemName)!;
        const requestsSheet = ensurePortalSheet_(PORTAL_REQUESTS_SHEET, PORTAL_REQUEST_HEADERS);
        const existingRequests = readPortalRequests_();
        const nextId = existingRequests.reduce((highest, existing) => Math.max(highest, existing.id), 0) + 1;
        const total = Number((item.unitPrice * quantity).toFixed(2));

        requestsSheet.appendRow([nextId, new Date(), identity.email, name, period, itemName, quarter, item.unitPrice, quantity, total, "Pending", "", "", onBehalf ? `Requested by ${identity.email} for the student` : ""]);
        if (onBehalf) log(`Portal: ${identity.email} requested ${quantity} x ${itemName} for ${name} (${period}).`, false);
        return { ok: true, message: `Requested ${quantity} x ${itemName} ($${total.toFixed(2)}). It is charged only if it is approved.` };
    } catch (error: any) {
        log(`Error in portalRequestExpenditure: ${error.message}`, false);
        return failure_("The request could not be sent. Please try again.");
    }
}

/**
 * A student takes back a request that is still pending.
 * @param data JSON string: { name, period, id }
 */
function portalCancelRequest(data: string): PortalActionResult {
    return runWithTransactionLock(() => {
        try {
            const request = JSON.parse(data);
            const name = String(request.name ?? "");
            const period = String(request.period ?? "");
            const requestId = Number(request.id);

            const { identity, reason, onBehalf } = authorizeProfileAction_(name, period);
            if (reason) return failure_(reason);

            const record = readPortalRequests_().find(candidate => candidate.id === requestId);
            if (!record || record.student !== name || record.period !== period) return failure_("That request was not found.");
            if (record.status !== "Pending") return failure_(`That request is already ${record.status.toLowerCase()}.`);

            const rowNumber = findPortalRequestRow_(requestId);
            getSheetNamed(PORTAL_REQUESTS_SHEET)!.getRange(rowNumber, 11, 1, 3).setValues([["Cancelled", onBehalf ? identity.email : name, new Date()]]);
            if (onBehalf) log(`Portal: ${identity.email} cancelled request ${requestId} for ${name} (${period}).`, false);
            return { ok: true, message: "Request cancelled." };
        } catch (error: any) {
            log(`Error in portalCancelRequest: ${error.message}`, false);
            return failure_("The request could not be cancelled. Please try again.");
        }
    }, failure_("The bank is busy right now. Please try again in a moment."));
}

// ---------------------------------------------------------------------------------------------------------------------------
// Admin actions
// ---------------------------------------------------------------------------------------------------------------------------

/** All requests, pending first. Admins only. */
function portalAdminListRequests(): any {
    try {
        const identity = resolvePortalIdentity_();
        if (identity.status !== "ok" || !identity.isAdmin) return { error: "Admins only." };

        const requests = readPortalRequests_();
        const pending = requests.filter(request => request.status === "Pending");
        const resolved = requests.filter(request => request.status !== "Pending").slice(-30).reverse();
        return { requests: [...pending, ...resolved] };
    } catch (error: any) {
        log(`Error in portalAdminListRequests: ${error.message}`, false);
        return { error: "The requests could not load." };
    }
}

/**
 * Approves or denies a request. Approving charges the student through the same code path as the Services dialog,
 * so the purchase is logged, counted against the caps and can be undone like any other.
 * @param data JSON string: { id: number, approve: boolean, note?: string }
 */
function portalAdminResolveRequest(data: string): PortalActionResult {
    return runWithTransactionLock(() => {
        try {
            const identity = resolvePortalIdentity_();
            if (identity.status !== "ok" || !identity.isAdmin) return failure_("Admins only.");

            const request = JSON.parse(data);
            const requestId = Number(request.id);
            const approve = request.approve === true;
            const note = String(request.note ?? "").slice(0, 200);

            const record = readPortalRequests_().find(candidate => candidate.id === requestId);
            if (!record) return failure_("That request was not found.");
            if (record.status !== "Pending") return failure_(`That request is already ${record.status.toLowerCase()}.`);

            const rowNumber = findPortalRequestRow_(requestId);
            const requestsSheet = getSheetNamed(PORTAL_REQUESTS_SHEET)!;

            if (!approve) {
                requestsSheet.getRange(rowNumber, 11, 1, 4).setValues([["Denied", identity.email, new Date(), note]]);
                return { ok: true, message: "Request denied." };
            }

            const student = findPortalStudent_(record.student, record.period);
            if (!student) return failure_(`${record.student} was not found on ${record.period}.`);

            // The price, balance and caps can have changed since the student asked.
            const refusal = checkPortalPurchase_(student, record.period, record.itemName, record.quarter as QuarterId, record.quantity, record.id);
            if (refusal) return failure_(`Can't approve: ${refusal}`);

            const currentItem = buildPortalCatalog_(record.quarter as QuarterId, record.student, record.period).find(candidate => candidate.itemName === record.itemName)!;
            const expendituresColumn = getExpendituresColumn();
            const periodSheet = getSheetNamed(record.period)!;

            const { result, messages } = runCapturingMessages_(() => applyMathToSelection(
                "ADD",
                currentItem.unitPrice,
                record.quantity,
                false,
                `$${currentItem.unitPrice} - ${record.itemName} (x${record.quantity})`,
                periodSheet.getRange(student.row, expendituresColumn),
                student.row,
                expendituresColumn
            ));

            if (result !== true) return failure_(messages.join(" ") || (typeof result === "string" ? result : "The charge could not be applied."));

            requestsSheet.getRange(rowNumber, 8, 1, 7).setValues([[currentItem.unitPrice, record.quantity, Number((currentItem.unitPrice * record.quantity).toFixed(2)), "Approved", identity.email, new Date(), note]]);
            return { ok: true, message: `Approved. ${record.student} was charged $${(currentItem.unitPrice * record.quantity).toFixed(2)}.` };
        } catch (error: any) {
            log(`Error in portalAdminResolveRequest: ${error.message}`, false);
            return failure_("The request could not be resolved. Please try again.");
        }
    }, failure_("The bank is busy right now. Please try again in a moment."));
}
