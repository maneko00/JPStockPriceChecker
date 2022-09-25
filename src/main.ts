import { JpStockGetter } from "./JpStockGetter";

export function myFunction() {
  let active_sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // アクティブな範囲の配列を取得
  var rangeList = active_sheet.getActiveRangeList().getRanges();
  // 取得したアクティブな範囲をループで回す
  for(var i = 0; i < rangeList.length ; i++){
    //アクティブセルがB列か判定
    let active_cell_column = rangeList[i].getColumn();
    const target_column = 2;
    if(active_cell_column != target_column)
    {
      continue;
    }

    let active_cell_row = rangeList[i].getRow();
    let active_cell_last_row = rangeList[i].getLastRow();
    for(var i = active_cell_row; i < active_cell_last_row+1 ; i++){
      if(i < 10)
      {
        continue;
      }
      if(i > 150)
      {
        continue;
      }

      let value = active_sheet.getRange(i, target_column).getValue();
      if(value == '')
      {
        active_sheet.getRange(i,3).setValue('');
        active_sheet.getRange(i,4).setValue('');
        active_sheet.getRange(i,6).setValue('');
        continue;
      }

      if(Number(value) == 'NaN')
      {
        continue;
      }
      let instance = new JpStockGetter(value);
      active_sheet.getRange(i,3).setValue(instance.name);
      active_sheet.getRange(i,4).setValue(instance.dividends_per_share);
      active_sheet.getRange(i,6).setValue(instance.price);
    }
  }
}
