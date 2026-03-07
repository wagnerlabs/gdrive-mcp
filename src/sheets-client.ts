import { google, sheets_v4 } from "googleapis";
import { OAuth2Client } from "google-auth-library";
import { handleApiError, DriveAPIError } from "./client.js";

export interface FormatOptions {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  fontSize?: number;
  fontFamily?: string;
  textColor?: string;
  backgroundColor?: string;
  horizontalAlignment?: "LEFT" | "CENTER" | "RIGHT";
  verticalAlignment?: "TOP" | "MIDDLE" | "BOTTOM";
  wrapStrategy?: "OVERFLOW_CELL" | "CLIP" | "WRAP";
  numberFormatType?: string;
  numberFormatPattern?: string;
}

function colToIndex(col: string): number {
  let index = 0;
  for (const ch of col.toUpperCase()) {
    index = index * 26 + (ch.charCodeAt(0) - 64);
  }
  return index - 1;
}

export class SheetsClient {
  private sheets: sheets_v4.Sheets;

  constructor(auth: OAuth2Client) {
    this.sheets = google.sheets({ version: "v4", auth });
  }

  async getSpreadsheet(spreadsheetId: string) {
    try {
      const res = await this.sheets.spreadsheets.get({
        spreadsheetId,
        fields:
          "spreadsheetId,properties,spreadsheetUrl,sheets.properties,namedRanges",
      });
      return {
        spreadsheetId: res.data.spreadsheetId!,
        title: res.data.properties?.title!,
        spreadsheetUrl: res.data.spreadsheetUrl!,
        locale: res.data.properties?.locale ?? undefined,
        sheets: (res.data.sheets ?? []).map((s) => ({
          sheetId: s.properties?.sheetId!,
          title: s.properties?.title!,
          index: s.properties?.index!,
          rowCount: s.properties?.gridProperties?.rowCount!,
          columnCount: s.properties?.gridProperties?.columnCount!,
        })),
        namedRanges: res.data.namedRanges ?? [],
      };
    } catch (err) {
      handleApiError(err, "Google Sheets");
    }
  }

  async createSpreadsheet(title: string, sheetNames?: string[]) {
    try {
      const sheets = (sheetNames ?? ["Sheet1"]).map((name, i) => ({
        properties: { title: name, index: i },
      }));
      const res = await this.sheets.spreadsheets.create({
        requestBody: { properties: { title }, sheets },
      });
      return {
        spreadsheetId: res.data.spreadsheetId!,
        spreadsheetUrl: res.data.spreadsheetUrl!,
        sheets: (res.data.sheets ?? []).map((s) => ({
          sheetId: s.properties?.sheetId!,
          title: s.properties?.title!,
          index: s.properties?.index!,
        })),
      };
    } catch (err) {
      handleApiError(err, "Google Sheets");
    }
  }

  async getValues(
    spreadsheetId: string,
    range: string,
  ): Promise<unknown[][]> {
    try {
      const res = await this.sheets.spreadsheets.values.get({
        spreadsheetId,
        range,
      });
      return res.data.values ?? [];
    } catch (err) {
      handleApiError(err, "Google Sheets");
    }
  }

  async updateValues(
    spreadsheetId: string,
    range: string,
    values: (string | number | boolean)[][],
    valueInputOption: string = "USER_ENTERED",
  ) {
    try {
      const res = await this.sheets.spreadsheets.values.update({
        spreadsheetId,
        range,
        valueInputOption,
        requestBody: { values },
      });
      return {
        spreadsheetId: res.data.spreadsheetId!,
        updatedRange: res.data.updatedRange!,
        updatedRows: res.data.updatedRows!,
        updatedColumns: res.data.updatedColumns!,
        updatedCells: res.data.updatedCells!,
      };
    } catch (err) {
      handleApiError(err, "Google Sheets");
    }
  }

  async appendValues(
    spreadsheetId: string,
    range: string,
    values: (string | number | boolean)[][],
    valueInputOption: string = "USER_ENTERED",
  ) {
    try {
      const res = await this.sheets.spreadsheets.values.append({
        spreadsheetId,
        range,
        valueInputOption,
        insertDataOption: "INSERT_ROWS",
        requestBody: { values },
      });
      return {
        spreadsheetId: res.data.spreadsheetId!,
        tableRange: res.data.tableRange ?? undefined,
        updatedRange: res.data.updates?.updatedRange!,
        updatedRows: res.data.updates?.updatedRows!,
        updatedCells: res.data.updates?.updatedCells!,
      };
    } catch (err) {
      handleApiError(err, "Google Sheets");
    }
  }

  async clearValues(spreadsheetId: string, range: string) {
    try {
      const res = await this.sheets.spreadsheets.values.clear({
        spreadsheetId,
        range,
        requestBody: {},
      });
      return {
        spreadsheetId: res.data.spreadsheetId!,
        clearedRange: res.data.clearedRange!,
      };
    } catch (err) {
      handleApiError(err, "Google Sheets");
    }
  }

  async formatCells(
    spreadsheetId: string,
    range: string,
    options: FormatOptions,
  ): Promise<void> {
    const parsed = SheetsClient.parseA1Range(range);
    const sheetId = await this.getSheetId(spreadsheetId, parsed.sheetTitle);

    const cellFormat: sheets_v4.Schema$CellFormat = {};
    const fields: string[] = [];

    const textFormat: sheets_v4.Schema$TextFormat = {};
    let hasTextFormat = false;

    if (options.bold !== undefined) {
      textFormat.bold = options.bold;
      fields.push("textFormat.bold");
      hasTextFormat = true;
    }
    if (options.italic !== undefined) {
      textFormat.italic = options.italic;
      fields.push("textFormat.italic");
      hasTextFormat = true;
    }
    if (options.underline !== undefined) {
      textFormat.underline = options.underline;
      fields.push("textFormat.underline");
      hasTextFormat = true;
    }
    if (options.strikethrough !== undefined) {
      textFormat.strikethrough = options.strikethrough;
      fields.push("textFormat.strikethrough");
      hasTextFormat = true;
    }
    if (options.fontSize !== undefined) {
      textFormat.fontSize = options.fontSize;
      fields.push("textFormat.fontSize");
      hasTextFormat = true;
    }
    if (options.fontFamily !== undefined) {
      textFormat.fontFamily = options.fontFamily;
      fields.push("textFormat.fontFamily");
      hasTextFormat = true;
    }
    if (options.textColor !== undefined) {
      textFormat.foregroundColor = SheetsClient.hexToColor(options.textColor);
      fields.push("textFormat.foregroundColor");
      hasTextFormat = true;
    }

    if (hasTextFormat) {
      cellFormat.textFormat = textFormat;
    }

    if (options.backgroundColor !== undefined) {
      cellFormat.backgroundColor = SheetsClient.hexToColor(
        options.backgroundColor,
      );
      fields.push("backgroundColor");
    }

    if (options.horizontalAlignment !== undefined) {
      cellFormat.horizontalAlignment = options.horizontalAlignment;
      fields.push("horizontalAlignment");
    }
    if (options.verticalAlignment !== undefined) {
      cellFormat.verticalAlignment = options.verticalAlignment;
      fields.push("verticalAlignment");
    }

    if (options.wrapStrategy !== undefined) {
      cellFormat.wrapStrategy = options.wrapStrategy;
      fields.push("wrapStrategy");
    }

    if (
      options.numberFormatType !== undefined ||
      options.numberFormatPattern !== undefined
    ) {
      cellFormat.numberFormat = {};
      if (options.numberFormatType !== undefined) {
        cellFormat.numberFormat.type = options.numberFormatType;
      }
      if (options.numberFormatPattern !== undefined) {
        cellFormat.numberFormat.pattern = options.numberFormatPattern;
      }
      fields.push("numberFormat");
    }

    try {
      await this.sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [
            {
              repeatCell: {
                range: {
                  sheetId,
                  startRowIndex: parsed.startRow,
                  startColumnIndex: parsed.startCol,
                  endRowIndex: parsed.endRow + 1,
                  endColumnIndex: parsed.endCol + 1,
                },
                cell: { userEnteredFormat: cellFormat },
                fields: fields
                  .map((f) => `userEnteredFormat.${f}`)
                  .join(","),
              },
            },
          ],
        },
      });
    } catch (err) {
      handleApiError(err, "Google Sheets");
    }
  }

  async addSheet(spreadsheetId: string, title: string) {
    try {
      const res = await this.sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [{ addSheet: { properties: { title } } }],
        },
      });
      const reply = res.data.replies?.[0]?.addSheet?.properties;
      return {
        sheetId: reply?.sheetId!,
        title: reply?.title!,
        index: reply?.index!,
      };
    } catch (err) {
      handleApiError(err, "Google Sheets");
    }
  }

  async deleteSheet(spreadsheetId: string, sheetTitle: string): Promise<void> {
    const sheetId = await this.getSheetId(spreadsheetId, sheetTitle);
    try {
      await this.sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [{ deleteSheet: { sheetId } }],
        },
      });
    } catch (err) {
      handleApiError(err, "Google Sheets");
    }
  }

  async renameSheet(
    spreadsheetId: string,
    currentTitle: string,
    newTitle: string,
  ) {
    const info = await this.getSpreadsheet(spreadsheetId);
    const sheet = info.sheets.find((s) => s.title === currentTitle);
    if (!sheet) {
      throw new DriveAPIError(
        `Sheet tab "${currentTitle}" not found in spreadsheet`,
      );
    }
    try {
      await this.sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [
            {
              updateSheetProperties: {
                properties: { sheetId: sheet.sheetId, title: newTitle },
                fields: "title",
              },
            },
          ],
        },
      });
    } catch (err) {
      handleApiError(err, "Google Sheets");
    }
    return { sheetId: sheet.sheetId, title: newTitle, index: sheet.index };
  }

  async insertDimension(
    spreadsheetId: string,
    sheetTitle: string,
    dimension: "ROWS" | "COLUMNS",
    startIndex: number,
    count: number,
  ): Promise<void> {
    const sheetId = await this.getSheetId(spreadsheetId, sheetTitle);
    try {
      await this.sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [
            {
              insertDimension: {
                range: {
                  sheetId,
                  dimension,
                  startIndex,
                  endIndex: startIndex + count,
                },
                inheritFromBefore: startIndex > 0,
              },
            },
          ],
        },
      });
    } catch (err) {
      handleApiError(err, "Google Sheets");
    }
  }

  async deleteDimension(
    spreadsheetId: string,
    sheetTitle: string,
    dimension: "ROWS" | "COLUMNS",
    startIndex: number,
    count: number,
  ): Promise<void> {
    const sheetId = await this.getSheetId(spreadsheetId, sheetTitle);
    try {
      await this.sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [
            {
              deleteDimension: {
                range: {
                  sheetId,
                  dimension,
                  startIndex,
                  endIndex: startIndex + count,
                },
              },
            },
          ],
        },
      });
    } catch (err) {
      handleApiError(err, "Google Sheets");
    }
  }

  async getSheetId(
    spreadsheetId: string,
    sheetTitle: string,
  ): Promise<number> {
    const info = await this.getSpreadsheet(spreadsheetId);
    const sheet = info.sheets.find((s) => s.title === sheetTitle);
    if (!sheet) {
      throw new DriveAPIError(
        `Sheet tab "${sheetTitle}" not found in spreadsheet`,
      );
    }
    return sheet.sheetId;
  }

  static parseA1Range(range: string): {
    sheetTitle: string;
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
  } {
    const match = range.match(
      /^(?:'([^']+)'|([^!]+))!([A-Za-z]+)(\d+)(?::([A-Za-z]+)(\d+))?$/,
    );
    if (!match) {
      throw new DriveAPIError(
        `Invalid range "${range}". Expected bounded A1 notation with sheet name, e.g. "Sheet1!A1:C5".`,
      );
    }
    const sheetTitle = match[1] ?? match[2];
    const startCol = colToIndex(match[3]);
    const startRow = parseInt(match[4], 10) - 1;
    const endCol = match[5] ? colToIndex(match[5]) : startCol;
    const endRow = match[6] ? parseInt(match[6], 10) - 1 : startRow;
    return { sheetTitle, startRow, startCol, endRow, endCol };
  }

  static hexToColor(hex: string): {
    red: number;
    green: number;
    blue: number;
  } {
    const h = hex.replace(/^#/, "");
    return {
      red: parseInt(h.substring(0, 2), 16) / 255,
      green: parseInt(h.substring(2, 4), 16) / 255,
      blue: parseInt(h.substring(4, 6), 16) / 255,
    };
  }
}
