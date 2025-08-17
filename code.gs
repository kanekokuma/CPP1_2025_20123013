//このスプレッドシートを操作の対象として設定
const sheetName = "シート1"; 

//スプレッドシートを取得する
function getSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    const message = `エラー: スプレッドシートに「${sheetName}」という名前のシートが見つかりません。シート名を確認してください。`;
    Logger.log(message);
    throw new Error(message);
  }
  return sheet;
}

//Webページ(index.html)を表示する
function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate();
}

//タスクデータをすべて取得する関数
function getToDos() {
  try {
    const sheet = getSheet();
    
    // ヘッダー行とデータ行の数を確認
    if (sheet.getLastRow() < 2) {
      return { success: true, data: [] }; // データ行がなければ空の配列を返す
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    
    const todoList = data.map(row => {
      const item = {};
      headers.forEach((header, index) => {
        switch(header) {
          case 'id':
            item.id = row[index];
            break;
          case 'todo':
            item.todo = row[index];
            break;
          case 'name':
            item.name = row[index];
            break;
          case 'date':
            item.date = row[index] ? new Date(row[index]).toLocaleDateString('ja-JP', { year: 'numeric', month: '2-digit', day: '2-digit' }).replace(/\//g, '-') : '';
            break;
          case 'status':
            item.status = row[index] || "未着手"; // ステータスが空の場合は「未着手」をデフォルトとして設定
            break;
          default:
            item[header] = row[index];
        }
      });
      return item;
    });
    Logger.log("getToDos: ToDoデータを正常に取得しました。");
    return { success: true, data: todoList };
  } catch (e) {
    Logger.log(`getToDos関数でエラーが発生しました: ${e.message}`);
    return { success: false, message: `ToDoリストの取得中にエラーが発生しました: ${e.message}` };
  }
}

//新しいToDoを追加する関数
function addTodo(todoData) {
  try {
    const sheet = getSheet();
    const newId = Utilities.getUuid();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    const newRow = headers.map(header => {
      switch(header) {
        case 'id': return newId;
        case 'todo': return todoData.todo || '';
        case 'name': return todoData.name || '';
        case 'date': return todoData.date || '';
        case 'status': return "未着手"; // 新規追加時は「未着手」をデフォルトで設定
        default: return '';
      }
    });

    sheet.appendRow(newRow);
    Logger.log(`addTodo: 「${todoData.todo}」を追加しました。ID: ${newId}`);
    return { success: true, message: `「${todoData.todo}」を追加しました。` };
  } catch (e) {
    Logger.log(`addTodo関数でエラーが発生しました: ${e.message}`);
    return { success: false, message: `ToDoの追加中にエラーが発生しました: ${e.message}` };
  }
}


//タスクを更新する関数
function updateTodo(item) {
  try {
    const sheet = getSheet();
    const idColumnIndex = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].indexOf('id');
    
    if (idColumnIndex === -1) {
      throw new Error("ヘッダーに 'id' 列が見つかりません。");
    }
    
    // TextFinderを使用してIDを検索
    const cell = sheet.getRange(2, idColumnIndex + 1, sheet.getLastRow() - 1, 1).createTextFinder(String(item.id)).matchEntireCell(true).findNext();
    
    if (cell) {
      const rowNum = cell.getRow();
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      
      const todoCol = headers.indexOf('todo') + 1;
      const nameCol = headers.indexOf('name') + 1;
      const dateCol = headers.indexOf('date') + 1;
      const statusCol = headers.indexOf('status') + 1;

      // すべての更新データを配列に格納
      const updateValues = [];
      const colMap = {
        'todo': todoCol,
        'name': nameCol,
        'date': dateCol,
        'status': statusCol
      };
      
      for (const key in item) {
        if (colMap[key] > 0) {
          updateValues.push({ col: colMap[key], value: item[key] });
        }
      }

      // 範囲をまとめて更新
      for (const update of updateValues) {
        sheet.getRange(rowNum, update.col).setValue(update.value);
      }
      
      Logger.log(`updateTodo: ID ${item.id} のToDoを更新しました。`);
      return { success: true, message: `ID: ${item.id} のToDoを更新しました。` };
    }
    Logger.log(`updateTodo: 更新対象のToDo (ID: ${item.id}) が見つかりません。`);
    return { success: false, message: "エラー: 更新対象のToDoが見つかりません。" };
  } catch (e) {
    Logger.log(`updateTodo関数でエラーが発生しました: ${e.message}`);
    return { success: false, message: `ToDoの更新中にエラーが発生しました: ${e.message}` };
  }
}

//タスクを削除する関数
function deleteTodo(id) {
  try {
    const sheet = getSheet();
    const idColumnIndex = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].indexOf('id');

    if (idColumnIndex === -1) {
      throw new Error("ヘッダーに 'id' 列が見つかりません。");
    }

    // TextFinderを使用してIDを検索
    const cell = sheet.getRange(2, idColumnIndex + 1, sheet.getLastRow() - 1, 1).createTextFinder(String(id)).matchEntireCell(true).findNext();
    
    if (cell) {
      sheet.deleteRow(cell.getRow()); 
      Logger.log(`deleteTodo: ID ${id} のToDoを削除しました。`);
      return { success: true, message: `ID: ${id} のToDoを削除しました。` };
    }
    Logger.log(`deleteTodo: 削除対象のToDo (ID: ${id}) が見つかりません。`);
    return { success: false, message: "エラー: 対象のToDoが見つかりません。" };
  } catch (e) {
    Logger.log(`deleteTodo関数でエラーが発生しました: ${e.message}`);
    return { success: false, message: `サーバーエラーが発生しました: ${e.message}` };
  }
}