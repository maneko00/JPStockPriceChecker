import { JpStockGetter } from "./JpStockGetter";

const kTargetColumn: number             = 2; // B列
const kNameColumn: number               = 3; // C列
const kDividendsPerShareColumn: number  = 4; // D列
const kStockPriceColumn: number         = 6; // F列
const kTargetFirstRow: number           = 10;
const kTargetLastRow: number            = 150;
const kNan: string                      = 'NaN';

export function ChangeFunction() {
    // １つ目のシートを取得
    let active_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    // アクティブな範囲の配列を取得
    var rangeList = active_sheet.getActiveRangeList()?.getRanges();
    // 取得したアクティブな範囲をループで回す
    for(var i = 0; i < rangeList!.length ; i++){
        //アクティブセルがB列か判定
        let active_cell_column = rangeList![i].getColumn();
        if(active_cell_column != kTargetColumn)
        {
            continue;
        }

        let active_cell_row = rangeList![i].getRow();
        let active_cell_last_row = rangeList![i].getLastRow();
        for(var i = active_cell_row; i < active_cell_last_row+1 ; i++){
            if(i < kTargetFirstRow)
            {
                continue;
            }
            if(i > kTargetLastRow)
            {
                continue;
            }

            let value = active_sheet.getRange(i, kTargetColumn).getValue();
            if(value == '')
            {
                active_sheet.getRange(i,kNameColumn).setValue('');
                active_sheet.getRange(i,kDividendsPerShareColumn).setValue('');
                active_sheet.getRange(i,kStockPriceColumn).setValue('');
                continue;
            }

            if(Number(value) == kNan)
            {
                continue;
            }
            let instance = new JpStockGetter(value);
            active_sheet.getRange(i,kNameColumn).setValue(instance.name);
            active_sheet.getRange(i,kDividendsPerShareColumn).setValue(instance.dividends_per_share);
            active_sheet.getRange(i,kStockPriceColumn).setValue(instance.price);
        }
    }
}

export function UpdateStockPrice() {
    // １つ目のシートを取得
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    // "B10:B150"まで選択範囲を指定
    const range_size = kTargetLastRow - kTargetFirstRow + 1;
    let range = sheet.getRange(kTargetFirstRow, kTargetColumn, range_size);
  
    for (var i = 0; i < range_size; i++)
    {
        let cell = range.getCell(i+1,1);
        Logger.log(cell.getRow() + "行目を処理中");
        
        if(cell.getValue() == '')
        {
            continue;
        }
        if (Number(cell.getValue()) == kNan) {
            continue;
        }
    
        let instance = new JpStockGetter(cell.getValue());
        sheet.getRange(cell.getRow(), kNameColumn).setValue(instance.name);
        sheet.getRange(cell.getRow(), kDividendsPerShareColumn).setValue(instance.dividends_per_share);
        sheet.getRange(cell.getRow(), kStockPriceColumn).setValue(instance.price);
    }
    Browser.msgBox("株価情報を更新しました。");
}
