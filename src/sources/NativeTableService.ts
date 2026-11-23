import { TableOptions } from '@interfaces/CollectionOptions';
import {
  NativeSheetsService,
  NativeTableState
} from '@interfaces/NativeTableServiceInterfaces';
import {
  NativeTableGridRange
} from '@shared-types/NativeTableTypes';

export class NativeTableService {
  private static readonly DEFAULT_PRESET = {
    headerColor: '#3C78D8',
    firstBandColor: '#FFFFFF',
    secondBandColor: '#EAF2F8'
  };
  private readonly service: NativeSheetsService;
  private table: NativeTableState | null | undefined;
  private readonly name: string;
  private readonly rowColors: {
    headerColor: string;
    firstBandColor: string;
    secondBandColor: string;
  };

  constructor(
    private readonly spreadsheetId: string,
    private readonly sheetId: number,
    collectionName: string,
    options: TableOptions,
    service?: NativeSheetsService
  ) {
    if (options.preset !== undefined && options.preset !== 'default') {
      throw new Error(`Table preset '${options.preset}' is not supported.`);
    }

    const globalService = (globalThis as typeof globalThis & {
      Sheets?: NativeSheetsService;
    }).Sheets;
    this.service = service ?? globalService ?? this.missingService();
    this.name = options.name ?? this.createDefaultName(collectionName);
    this.rowColors = {
      headerColor: options.headerColor ?? NativeTableService.DEFAULT_PRESET.headerColor,
      firstBandColor: options.firstBandColor ?? NativeTableService.DEFAULT_PRESET.firstBandColor,
      secondBandColor: options.secondBandColor ?? NativeTableService.DEFAULT_PRESET.secondBandColor
    };

    Object.values(this.rowColors).forEach(color => this.validateColor(color));
  }

  findTable(refresh = false): NativeTableState | null {
    if (!refresh && this.table !== undefined) {
      return this.table;
    }

    const metadata = this.service.Spreadsheets.get(this.spreadsheetId, {
      fields: 'sheets(properties(sheetId,title),tables(tableId,name,range))'
    });
    const sheet = metadata.sheets?.find(
      candidate => candidate.properties?.sheetId === this.sheetId
    );
    const table = sheet?.tables?.find(candidate => candidate.name === this.name);

    if (table) {
      this.table = {
        tableId: table.tableId,
        name: table.name,
        range: table.range
      };
      return this.table;
    }

    const conflictingTable = sheet?.tables?.[0];

    if (conflictingTable) {
      throw new Error(
        `The tab already contains the native table '${conflictingTable.name}', ` +
        `but the collection expects '${this.name}'.`
      );
    }

    this.table = null;
    return null;
  }

  createTable(values: unknown[][]): NativeTableState {
    const headers = this.validateValues(values);

    const range: NativeTableGridRange = {
      sheetId: this.sheetId,
      startRowIndex: 0,
      endRowIndex: values.length,
      startColumnIndex: 0,
      endColumnIndex: headers.length
    };
    const response = this.service.Spreadsheets.batchUpdate({
      requests: [{
        addTable: {
          table: {
            name: this.name,
            range,
            columnProperties: headers.map((columnName, columnIndex) => ({
              columnName,
              columnIndex
            })),
            rowsProperties: {
              headerColorStyle: this.toColorStyle(this.rowColors.headerColor),
              firstBandColorStyle: this.toColorStyle(this.rowColors.firstBandColor),
              secondBandColorStyle: this.toColorStyle(this.rowColors.secondBandColor)
            }
          }
        }
      }]
    }, this.spreadsheetId);
    const created = response.replies?.[0]?.addTable?.table;

    if (!created?.tableId) {
      throw new Error('The Sheets API did not return the identifier for the created table.');
    }

    this.table = {
      tableId: created.tableId,
      name: created.name || this.name,
      range: created.range || range
    };

    return this.table;
  }

  validateValues(values: unknown[][]): string[] {
    const headers = values[0]?.map(String) ?? [];
    this.validateHeaders(headers);
    return headers;
  }

  resizeTable(table: NativeTableState, values: unknown[][]): void {
    const headers = values[0]?.map(String) ?? [];
    this.validateHeaders(headers);

    const range: NativeTableGridRange = {
      ...table.range,
      endRowIndex: table.range.startRowIndex + values.length,
      endColumnIndex: table.range.startColumnIndex + headers.length
    };
    this.service.Spreadsheets.batchUpdate({
      requests: [{
        updateTable: {
          table: {
            tableId: table.tableId,
            range,
            columnProperties: headers.map((columnName, columnIndex) => ({
              columnName,
              columnIndex
            }))
          },
          fields: 'range,columnProperties'
        }
      }]
    }, this.spreadsheetId);

    this.table = { ...table, range };
  }

  resizeTableRows(table: NativeTableState, rowCount: number): NativeTableState {
    const range: NativeTableGridRange = {
      ...table.range,
      endRowIndex: table.range.startRowIndex + rowCount
    };
    this.service.Spreadsheets.batchUpdate({
      requests: [{
        updateTable: {
          table: { tableId: table.tableId, range },
          fields: 'range'
        }
      }]
    }, this.spreadsheetId);
    this.table = { ...table, range };
    return this.table;
  }

  deleteTable(table: NativeTableState): void {
    this.service.Spreadsheets.batchUpdate({
      requests: [{ deleteTable: { tableId: table.tableId } }]
    }, this.spreadsheetId);
    this.table = null;
  }

  private validateHeaders(headers: string[]): void {
    if (headers.length === 0 || headers.some(header => header.trim() === '')) {
      throw new Error('Table headers cannot be empty.');
    }
    if (new Set(headers).size !== headers.length) {
      throw new Error('Table headers must be unique.');
    }
  }

  private validateColor(color: string): void {
    if (!/^#[0-9a-f]{6}$/i.test(color)) {
      throw new Error(`Invalid preset color '${color}'. Use hexadecimal format #RRGGBB.`);
    }
  }

  private toColorStyle(color: string): Record<string, unknown> {
    const normalized = color.slice(1);
    return {
      rgbColor: {
        red: parseInt(normalized.slice(0, 2), 16) / 255,
        green: parseInt(normalized.slice(2, 4), 16) / 255,
        blue: parseInt(normalized.slice(4, 6), 16) / 255
      }
    };
  }

  private createDefaultName(collectionName: string): string {
    const safeName = collectionName
      .replace(/[^a-zA-Z0-9_]/g, '_')
      .replace(/^[^a-zA-Z_]+/, '');
    return `SheetCollection_${safeName || 'Table'}`;
  }

  private missingService(): never {
    throw new Error(
      'The Advanced Sheets service is unavailable. Enable the advanced "Sheets API" service in the Apps Script project.'
    );
  }
}
