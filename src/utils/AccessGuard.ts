/* Coded by Erick Tran for Mr. Banderas, 2026
 * Copyright (c) 2026 Erick Tran.
 * This file is licensed under the MIT License, check the LICENSE file for details.
 *
 * GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
 *
 * !!!--- IMPORTANT NOTE: This file was originally written in TypeScript, go to the GitHub to see the original non-compiled code. ---!!!
 *
 * Who may run the admin, management and data functions.
 *
 * Every public function of an Apps Script project can be called from a web page of the same project with google.script.run, and the student
 * portal is such a page. So the functions that change money, settings, caches or sheets, or that read every student's data, start with
 * requireEditor_(): it lets the call through only when the person who is running it can edit this spreadsheet.
 *
 * Code that is allowed to use those functions on behalf of somebody who is NOT an editor (the portal's own endpoints, which check the
 * caller themselves, and the unattended entry points: triggers) runs inside runTrusted_(). Nothing a browser sends can turn that on.
 */

/** How many runTrusted_ calls are active right now. Starts at 0 for every request, because each server call is a fresh execution. */
let trustedCallDepth: number = 0;

/** Result of the editor check, remembered for the rest of this execution. */
let editorCheckResult: boolean | null = null;

/** Runs work with the editor check switched off, for code that has already decided the caller may do it (or that nobody calls directly). */
function runTrusted_<T>(work: () => T): T {
    trustedCallDepth++;
    try {
        return work();
    } finally {
        trustedCallDepth--;
    }
}

/**
 * Throws unless the person running this call can edit the spreadsheet (or the caller is trusted code).
 * The person is whoever Google says is running the script: the editor for dialogs and menus, the visitor for the web app. A visitor from
 * outside the owner's Google Workspace domain has no visible email, so they are never an editor.
 */
function requireEditor_(): void {
    if (!ENFORCE_EDITOR_GUARD || trustedCallDepth > 0) return;

    if (editorCheckResult === null) {
        const callerEmail = (Session.getActiveUser().getEmail() || "").trim().toLowerCase();
        editorCheckResult = callerEmail !== "" && getPortalAdminEmails_().includes(callerEmail);
        if (!editorCheckResult) log(`Blocked a call from ${callerEmail || "an account with no visible email"}: only people who can edit this spreadsheet may do that.`, false);
    }
    if (!editorCheckResult) throw new Error("Only people who can edit this spreadsheet can do that.");
}
