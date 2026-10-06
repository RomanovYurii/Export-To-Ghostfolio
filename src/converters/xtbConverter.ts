import * as fs from "fs";
import path from "path";
import dayjs from "dayjs";
import { parse } from "csv-parse";
import ExcelJS from "exceljs";
import { XtbRecord } from "../models/xtbRecord";
import { AbstractConverter } from "./abstractconverter";
import { SecurityService } from "../securityService";
import { GhostfolioExport } from "../models/ghostfolioExport";
import YahooFinanceRecord from "../models/yahooFinanceRecord";
import customParseFormat from "dayjs/plugin/customParseFormat";
import { GhostfolioOrderType } from "../models/ghostfolioOrderType";
import { getTags } from "../helpers/tagHelpers";

export class XtbConverter extends AbstractConverter {

    private static readonly CASH_OPERATIONS_SHEET = "Cash Operations";

    private static readonly REQUIRED_HEADER_COLUMNS = ["type", "ticker", "time", "amount", "id", "comment"];

    constructor(securityService: SecurityService) {
        super(securityService);

        dayjs.extend(customParseFormat);
    }

    /**
     * @inheritdoc
     */
    public readAndProcessFile(inputFile: string, successCallback: CallableFunction, errorCallback: CallableFunction): void {

        // XLSX exports are binary, so they need a dedicated reading path.
        if (path.extname(inputFile).toLocaleLowerCase() === ".xlsx") {

            if (!fs.existsSync(inputFile)) {
                return errorCallback(new Error(`File ${inputFile} does not exist!`));
            }

            return this.processWorkbook(inputFile, successCallback, errorCallback);
        }

        super.readAndProcessFile(inputFile, successCallback, errorCallback);
    }

    /**
     * @inheritdoc
     */
    public processFileContents(input: string, successCallback: any, errorCallback: any): void {

        // Parse the CSV and convert to Ghostfolio import format.
        parse(input, {
            delimiter: ";",
            fromLine: 2,
            skip_empty_lines: true,
            columns: this.processHeaders(input),
            cast: (columnValue, context) => {

                // Custom mapping below.

                // Convert type to Ghostfolio type.
                if (context.column === "type") {
                    return this.mapType(columnValue);
                }

                if (context.column === "symbol") {
                    return this.mapSymbol(columnValue);
                }

                // Parse numbers to floats (from string).
                if (context.column === "id" || context.column === "amount") {
                    return parseFloat(columnValue);
                }

                return columnValue;
            },
            on_record: (record: XtbRecord) => this.postProcessRecord(record)
        }, async (err, records: XtbRecord[]) => {

            try {

                // Check if parsing failed..
                if (err || records === undefined || records.length === 0) {
                    let errorMsg = "An error occurred while parsing!";

                    if (err) {
                        errorMsg += ` Details: ${err.message}`
                    }

                    return errorCallback(new Error(errorMsg))
                }

                console.log("[i] Read CSV file. Start processing..");

                await this.processRecords(records, successCallback, errorCallback);
            }
            catch (error) {
                this.handleProcessingError(error, errorCallback);
            }
        });
    }

    /**
     * Read and process an XLSX export file.
     *
     * @param inputFile The XLSX file to convert.
     * @param successCallback A callback to execute after processing has succeeded.
     * @param errorCallback A callback to execute after processing has failed.
     */
    private processWorkbook(inputFile: string, successCallback: CallableFunction, errorCallback: CallableFunction): void {

        (async () => {

            try {

                const workbook = new ExcelJS.Workbook();
                await workbook.xlsx.readFile(inputFile);

                const records = this.parseCashOperationsSheet(workbook);

                // Check if the sheet contained any records.
                if (records.length === 0) {
                    return errorCallback(new Error("An error occurred while parsing!"));
                }

                console.log("[i] Read XLSX file. Start processing..");

                await this.processRecords(records, successCallback, errorCallback);
            }
            catch (error) {
                this.handleProcessingError(error, errorCallback);
            }
        })();
    }

    /**
     * Log and pass on an error that occurred while processing the file contents.
     *
     * @param error The error that occurred.
     * @param errorCallback A callback to execute after processing has failed.
     */
    private handleProcessingError(error: any, errorCallback: CallableFunction): void {
        console.log("[e] An error occurred while processing the file contents. Stack trace:");
        console.log(error.stack);
        this.progress.stop();
        errorCallback(error);
    }

    /**
     * Parse the "Cash Operations" sheet of an XTB XLSX export into records.
     *
     * @param workbook The workbook to extract the records from.
     * @returns The records of the "Cash Operations" sheet.
     */
    private parseCashOperationsSheet(workbook: ExcelJS.Workbook): XtbRecord[] {

        const worksheet = workbook.getWorksheet(XtbConverter.CASH_OPERATIONS_SHEET);

        if (!worksheet) {
            throw new Error(`Sheet '${XtbConverter.CASH_OPERATIONS_SHEET}' not found in the export file!`);
        }

        // Locate the header row (row 5 in current exports) by looking for the known column names,
        // and map the header names to their column indexes.
        const columnMap = new Map<string, number>();
        let headerRowIndex = 0;

        for (let rowNumber = 1; rowNumber <= Math.min(worksheet.rowCount, 20); rowNumber++) {
            const cells = worksheet.getRow(rowNumber).values as ExcelJS.CellValue[];
            const headers = cells.map(cell => this.cellToString(cell).toLocaleLowerCase());

            if (XtbConverter.REQUIRED_HEADER_COLUMNS.every(column => headers.includes(column))) {
                headerRowIndex = rowNumber;

                cells.forEach((cell, index) => {
                    const header = this.cellToString(cell).toLocaleLowerCase();
                    if (header !== "") {
                        columnMap.set(header, index);
                    }
                });

                break;
            }
        }

        if (headerRowIndex === 0) {
            throw new Error(`Could not find the header row in the '${XtbConverter.CASH_OPERATIONS_SHEET}' sheet!`);
        }

        const records: XtbRecord[] = [];

        // Read the data rows until the trailing total (or the end of the sheet) is reached.
        for (let rowNumber = headerRowIndex + 1; rowNumber <= worksheet.rowCount; rowNumber++) {

            const row = worksheet.getRow(rowNumber);
            const type = this.cellToString(row.getCell(columnMap.get("type")).value);

            // Stop at the trailing total row.
            if (type.toLocaleLowerCase() === "total") {
                break;
            }

            // Skip rows without an operation type.
            if (type === "") {
                continue;
            }

            const record: XtbRecord = {
                id: Number(this.cellToString(row.getCell(columnMap.get("id")).value)),
                type: this.mapType(type),
                time: dayjs(this.cellToDate(row.getCell(columnMap.get("time")).value)).format("DD.MM.YYYY HH:mm:ss"),
                symbol: this.mapSymbol(this.cellToString(row.getCell(columnMap.get("ticker")).value)),
                comment: this.cellToString(row.getCell(columnMap.get("comment")).value),
                amount: this.cellToNumber(row.getCell(columnMap.get("amount")).value)
            };

            records.push(this.postProcessRecord(record));
        }

        return records;
    }

    /**
     * Convert the parsed records to a Ghostfolio import file.
     *
     * @param records The records to convert.
     * @param successCallback A callback to execute after processing has succeeded.
     * @param errorCallback A callback to execute after processing has failed.
     */
    private async processRecords(records: XtbRecord[], successCallback: CallableFunction, errorCallback: CallableFunction): Promise<void> {

        const result: GhostfolioExport = {
            meta: {
                date: new Date(),
                version: "v0"
            },
            activities: []
        }

        // Populate the progress bar.
        const bar1 = this.progress.create(records.length, 0);

        for (let idx = 0; idx < records.length; idx++) {
            const record = records[idx];

            // Check if the record should be ignored.
            if (this.isIgnoredRecord(record)) {
                bar1.increment();
                continue;
            }

            const date = dayjs(`${record.time}`, "DD.MM.YYYY HH:mm:ss");

            // Interest does not have a security, so add those immediately.
            if (record.type.toLocaleLowerCase() === "interest") {

                // Add interest record to export. The dataSource is omitted on
                // purpose: Ghostfolio resolves it to MANUAL for interest records
                // and rejects MANUAL activities with a free-text symbol.
                result.activities.push({
                    accountId: process.env.GHOSTFOLIO_ACCOUNT_ID,
                    comment: `XTB ${record.id} - ${record.comment}`,
                    fee: 0,
                    quantity: 1,
                    type: GhostfolioOrderType[record.type],
                    unitPrice: Math.abs(record.amount),
                    currency: process.env.XTB_ACCOUNT_CURRENCY || "EUR",
                    date: date.format("YYYY-MM-DDTHH:mm:ssZ"),
                    symbol: record.comment,
                    tags: getTags()
                });

                bar1.increment();
                continue;
            }

            if (record.type.toLocaleLowerCase() === "fee") {

                // Add fee record to export. The dataSource is omitted on
                // purpose, see the interest record above.
                result.activities.push({
                    accountId: process.env.GHOSTFOLIO_ACCOUNT_ID,
                    comment: `XTB ${record.id} - ${record.comment}`,
                    fee: Math.abs(record.amount),
                    quantity: 1,
                    type: GhostfolioOrderType[record.type],
                    unitPrice: 0,
                    currency: process.env.XTB_ACCOUNT_CURRENCY || "EUR",
                    date: date.format("YYYY-MM-DDTHH:mm:ssZ"),
                    symbol: record.comment,
                    tags: getTags()
                });

                bar1.increment();
                continue;
            }

            const match = record.comment.match(/(?:OPEN|CLOSE) BUY ([0-9]+(?:\.[0-9]+)?(?:\/[0-9]+(?:\.[0-9]+)?)?) @ ([0-9]+(?:\.[0-9]+)?)|(?:[A-Z\. ]+) ([0-9]+(?:\.[0-9]+)?)/)

            let quantity = parseFloat(match[1]?.split("/")[0]);
            let unitPrice = parseFloat(match[2]);
            const dividendPerShare = parseFloat(match[3]);

            // By default, there is no expected currency.
            let expectedCurrency = null;

            // For dividends and spin-offs, the currency is in the comment.
            if (record.type.toLocaleLowerCase() === "dividend") {
                expectedCurrency = record.comment.split(" ")[1];
            }

            let security: YahooFinanceRecord;
            try {
                security = await this.securityService.getSecurity(
                    null,
                    record.symbol,
                    null,
                    expectedCurrency,
                    this.progress);
            }
            catch (err) {
                this.logQueryError(record.comment, idx + 2);
                return errorCallback(err);
            }

            // Log whenever there was no match found.
            if (!security) {
                this.progress.log(`[i] No result found for action ${record.type}, symbol ${record.symbol} and comment ${record.comment}! Please add this manually..\n`);
                bar1.increment();
                continue;
            }

            let feeAmount = 0;

            // Dividend usually goes with a dividend tax record, so look it up.
            if (record.type.toLocaleLowerCase() === "dividend") {

                unitPrice = dividendPerShare;
                quantity = parseFloat((record.amount / dividendPerShare).toFixed(2));

                const taxRecord = this.lookupDividendTaxRecord(record.id, records, idx);

                // If there was a dividend tax record found, check if it matches the dividend record.
                if (taxRecord && taxRecord.symbol === record.symbol && taxRecord.time === record.time) {
                    feeAmount = Math.abs(taxRecord.amount);
                }
            }

            // Add record to export.
            result.activities.push({
                accountId: process.env.GHOSTFOLIO_ACCOUNT_ID,
                comment: `XTB ${record.id} - ${record.comment}`,
                fee: feeAmount,
                quantity: quantity,
                type: GhostfolioOrderType[record.type],
                unitPrice: unitPrice,
                currency: security.currency,
                dataSource: "YAHOO",
                date: date.format("YYYY-MM-DDTHH:mm:ssZ"),
                symbol: security.symbol,
                tags: getTags()
            });

            bar1.increment();
        }

        this.progress.stop();

        successCallback(result);
    }

    /**
     * Map the raw operation type of a record to a Ghostfolio record type.
     *
     * @param rawType The raw operation type from the export.
     * @returns The mapped Ghostfolio record type, or the raw type when no mapping applies.
     */
    private mapType(rawType: string): string {

        const type = rawType.toLocaleLowerCase();

        if (type.indexOf("stock purchase") > -1 || type.indexOf("stocks/etf purchase") > -1 || type.indexOf("ações/etf compra") > -1) {
            return "buy";
        }
        else if (type.indexOf("stock sell") > -1 || type.indexOf("stocks/etf sale") > -1 || type.indexOf("ações/etf vende") > -1) {
            return "sell";
        }
        else if (type.indexOf("sec fee") > -1 || type.indexOf("swap") > -1 || type.indexOf("commission") > -1 || type.indexOf("free funds interest tax") > -1 || type.indexOf("free funds interests tax") > -1) {
            return "fee";
        }
        else if (type.indexOf("free funds interest") > -1) {
            return "interest";
        }
        else if (type.indexOf("dividend") > -1 || type.indexOf("spin off") > -1) { //verify spinoff
            return "dividend";
        }
        else if (type.indexOf("profit/loss") > -1 || type.indexOf("close trade") > -1) {
            return "profitloss";
        }

        return rawType;
    }

    /**
     * Map the raw symbol of a record to a Yahoo Finance symbol.
     *
     * @param rawSymbol The raw symbol from the export.
     * @returns The mapped Yahoo Finance symbol.
     */
    private mapSymbol(rawSymbol: string): string {

        // XTB uses the .UK postfix, whereas Yahoo Finance uses .L for the London Stock Exchange.
        if (rawSymbol.endsWith(".UK")) {
            return rawSymbol.replace(".UK", ".L");
        }

        return rawSymbol;
    }

    /**
     * Apply post-processing on a parsed record.
     *
     * @param record The record to post-process.
     * @returns The post-processed record.
     */
    private postProcessRecord(record: XtbRecord): XtbRecord {

        // If a record is typed as dividend, but is a negative amount, then change type to fee.
        if (record.type === "dividend" && record.amount < 0) {
            record.type = "fee";
        }

        // If a record is typed as interest, but is a negative correction, then change type to fee.
        if (record.type === "interest" && record.comment.toLocaleLowerCase().startsWith("corr")) {
            record.type = record.amount < 0 ? "fee" : "interest";
        }

        // If the record is a profit/loss, check if it should be a fee or interest.
        if (record.type.toLocaleLowerCase() === "profitloss") {
            if (record.amount < 0) {
                record.type = "fee";
            }
            else {
                record.type = "interest";
            }
        }

        return record;
    }

    /**
     * Convert an XLSX cell value to a string.
     *
     * @param value The cell value to convert.
     * @returns The string representation of the cell value.
     */
    private cellToString(value: ExcelJS.CellValue): string {

        if (value === null || value === undefined) {
            return "";
        }

        if (value instanceof Date) {
            return value.toISOString();
        }

        if (typeof value === "object") {
            // Fall back to the text of rich text or hyperlink cells.
            if ("richText" in value) {
                return value.richText.map(part => part.text).join("");
            }
            if ("text" in value) {
                return `${value.text}`;
            }
            if ("result" in value) {
                return `${value.result}`;
            }
        }

        return `${value}`.trim();
    }

    /**
     * Convert an XLSX cell value to a number.
     *
     * @param value The cell value to convert.
     * @returns The numeric representation of the cell value.
     */
    private cellToNumber(value: ExcelJS.CellValue): number {

        if (typeof value === "number") {
            return value;
        }

        return parseFloat(this.cellToString(value));
    }

    /**
     * Convert an XLSX cell value to a date.
     *
     * @param value The cell value to convert.
     * @returns The date representation of the cell value.
     */
    private cellToDate(value: ExcelJS.CellValue): Date {

        if (value instanceof Date) {
            return value;
        }

        if (typeof value === "number") {
            // Excel stores dates as a serial number of days since 01-01-1970 (minus the leap bug days).
            return new Date(Math.round((value - 25569) * 86400000));
        }

        const parsed = dayjs(this.cellToString(value), "DD.MM.YYYY HH:mm:ss");

        return parsed.isValid() ? parsed.toDate() : new Date(this.cellToString(value));
    }

    /**
     * @inheritdoc
     */
    protected processHeaders(_: string): string[] {

        // Generic header mapping from the XTB CSV export.
        const csvHeaders = [
            "id",
            "type",
            "time",
            "symbol",
            "comment",
            "amount"];

        return csvHeaders;
    }

    /**
     * @inheritdoc
     */
    public isIgnoredRecord(record: XtbRecord): boolean {
        let ignoredRecordTypes = ["deposit", "withdrawal", "tax", "transfer"];

        return ignoredRecordTypes.some(t => record.type.toLocaleLowerCase().indexOf(t) > -1)
    }

    private lookupDividendTaxRecord(currentRecordId: number, records: XtbRecord[], idx: number): XtbRecord | undefined {

        let taxRecord;

        // Look ahead at the next record (if there are any records left).
        if (idx > 0 && records.length - 1 > idx + 1) {
            const nextRecord = records[idx + 1];
            if (nextRecord.type.toLocaleLowerCase().indexOf("tax") > -1 && currentRecordId + 1 === nextRecord.id) {
                taxRecord = nextRecord;
            }
        }

        // If there is no tax record found, look back at the previous record.
        if (!taxRecord) {

            const previousRecord = records[idx - 1];
            if (previousRecord?.type.toLocaleLowerCase().indexOf("tax") > -1 && currentRecordId + 1 === previousRecord?.id) {
                taxRecord = previousRecord;
            }
        }

        return taxRecord;
    }
}
