// Coded by Erick Tran for Mr. Banderas, 2026
// Copyright (c) 2026 Erick Tran. 
// This file is licensed under the MIT License, check the LICENSE file for details.

// GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
// This file contains utility functions for various operations in the B-Bucks Economy Scripts project, including functions related to services management and other financial operations.

interface QuarterlyData {
    Q1?: number;
    Q2?: number;
    Q3?: number;
    Q4?: number;
}

interface ItemData {
    itemName: string;
    category: "Income" | "Expense";
    pricing: QuarterlyData;
    limit: QuarterlyData;
}

function fetchServicesData(): ItemData[] | { error: string } {
    try {
        const SHEET_NAME = "Services";
        const ROW_START = 3;

        const columns: { [key: string]: number | number[] } = {
            itemName: 1,
            category: 2,
            // Represents Q1-Q4 pricing columns
            pricing: [3, 4, 5, 6],
            // Represents Q1-Q4 max per person columns
            limit: [7, 8, 9, 10]
        }

        const servicesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
        if (!servicesSheet) {
            Logger.log("Services sheet not found.");
            SpreadsheetApp.getUi().alert("Services sheet not found.");
            return { error: "Services sheet not found." };
        }

        const lastRow = servicesSheet.getLastRow();
        if (lastRow < ROW_START) {
            Logger.log("No services data found.");
            SpreadsheetApp.getUi().alert("No services data found.");
            return { error: "No services data found." };
        }

        const servicesData: ItemData[] = [];

        // Extract data from each row
        for (let row = ROW_START; row <= lastRow; row++) {
            const itemName = servicesSheet.getRange(row, columns.itemName as number).getValue().toString().trim();
            const category = servicesSheet.getRange(row, columns.category as number).getValue().toString().trim() as "Income" | "Expense";
            const pricing: QuarterlyData = {};
            const limit: QuarterlyData = {};

            // Extract pricing for Q1-Q4
            (columns.pricing as number[]).forEach((col, index) => {
                const value = servicesSheet.getRange(row, col).getValue();
                if (value !== "") {
                    pricing[`Q${index + 1}` as keyof QuarterlyData] = Number(value);
                }
            });

            // Extract limits for Q1-Q4
            (columns.limit as number[]).forEach((col, index) => {
                const value = servicesSheet.getRange(row, col).getValue();
                if (value !== "") {
                    limit[`Q${index + 1}` as keyof QuarterlyData] = Number(value);
                }
            });

            // Only include items that have a name, category, and at least one pricing
            if (itemName && category && Object.keys(pricing).length > 0) {
                servicesData.push({
                    itemName,
                    category,
                    pricing,
                    limit
                });
            }
        }

        if (servicesData.length === 0) {
            Logger.log("No valid services data found.");
            SpreadsheetApp.getUi().alert("No valid services data found.");
            return { error: "No valid services data found." };
        }

        return servicesData;
    } catch (error: any) {
        SpreadsheetApp.getUi().alert(`Error occurred in fetchServicesData: ${error.message}`);
        return { error: `Error occurred in fetchServicesData: ${error.message}` };
    }
}