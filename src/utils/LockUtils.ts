/* Coded by Erick Tran for Mr. Banderas, 2026
 * Copyright (c) 2026 Erick Tran.
 * This file is licensed under the MIT License, check the LICENSE file for details.
 *
 * GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
 *
 * !!!--- IMPORTANT NOTE: This file was originally written in TypeScript, go to the GitHub to see the original non-compiled code. ---!!!
 *
 * One lock for everything that moves money or appends to the transaction log.
 * Without it, two admins acting at the same moment can read the same cell, both write, and one change is lost; or both pass a
 * Max/Person check; or both append transactions with the same id.
 *
 * The DOCUMENT lock is used on purpose: the caching code already takes the SCRIPT lock for itself, and keeping the two apart means a
 * cache refresh in the middle of a transaction can never release the transaction's lock early.
 */

// How long a transaction waits for another one to finish before giving up.
const TRANSACTION_LOCK_WAIT_MS: number = 30000;

// Nested calls (for example undo -> addTransactionRecords) reuse the lock this execution already holds.
let transactionLockDepth: number = 0;

/**
 * Runs `work` while holding the transaction lock.
 * @param work The code to run.
 * @param valueIfBusy Returned when the lock could not be obtained in time (nothing has been changed in that case).
 */
function runWithTransactionLock<T>(work: () => T, valueIfBusy: T): T {
    if (transactionLockDepth > 0) {
        transactionLockDepth++;
        try {
            return work();
        } finally {
            transactionLockDepth--;
        }
    }

    const lock = LockService.getDocumentLock() || LockService.getScriptLock();
    if (!lock.tryLock(TRANSACTION_LOCK_WAIT_MS)) {
        log("Another transaction is still running. Nothing was changed, please try again in a moment.", true);
        return valueIfBusy;
    }

    transactionLockDepth = 1;
    try {
        return work();
    } finally {
        transactionLockDepth = 0;
        lock.releaseLock();
    }
}
