// イベント登録
function createEvent(calendarId, users) {
  // リクエストIDをランダムに生成
  const requestId = Utilities.formatString("shuffle#%d", Math.random()*100);
  // イベントの詳細設定
  const detail = {
    summary: 'シャッフルランチ',
    location: 'リモート',
    description: '同期との交流を目的としたランチ会です！',
    start: {
      dateTime: new Date('2020/8/3 12:00:00').toISOString()
    },
    end: {
      dateTime: new Date('2020/8/3 12:30:00').toISOString()
    },
    conferenceData: {
      createRequest: {
        conferenceSolutionKey: {
          type: "hangoutsMeet"
        },
        requestId: requestId
      }
    },
    attendees: users
  };
  // イベントの登録
  const event = Calendar.Events.insert(detail, calendarId, { conferenceDataVersion: 1 });
  Logger.log('Event ID: ' + JSON.stringify(event));
}

// 配列の要素をシャッフル
function shuffle(array){
  var result = [];
  for(i = array.length; i > 0; i--){
    var index = Math.floor(Math.random() * i);
    var val = array.splice(index, 1)[0];
    result.push(val);
  }

  return result;
}

// グループ作成
function createLunchGroups(array) {
  const arrayLen = array.length;
  // 作成できるグループ数
  const groupNum = Math.floor(arrayLen / 3);
  var result = [];

  // グループ数が2つ未満だった場合、グループ分けを行わない
  if (groupNum < 2) {
    result.push(array)
    return result;
  }
  // 3人一組でグループ分けを行う
  switch (arrayLen % 3) {
    // 1人余る場合: 4人グループを1つ作る
    case 1:
      // 4人グループを1つ作成
      result.push(array.slice(0, 4));
      // 3人グループを作成
      for (i = 2; i <= groupNum; i++) {
        result.push(array.slice(i*3-2, i*3+1));
      }
      break;
    // 2人余る場合
    case 2:
      // 4人グループを2つ作成
      for (i = 0; i < 2; i++) {
        result.push(array.slice(i*4, i*4+4))
      }
      // 3人グループ作成
      for (i = 3; i <= groupNum; i++) {
        result.push(array.slice(i*3-1, i*3+2));
      }
      break;
    // 3人ちょうどで作れる場合
    default:
      for (i = 1; i <= groupNum; i++) {
        result.push(array.slice(i*3-3, i*3))
      }
      break;
  }

  return result;
}

function mainProcess() {
  // シートオブジェクトを作成
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // 操作するシートのオブジェクトを取得
  const sheet = ss.getSheetByName('test');
  // ユーザのEmailアドレスをシートから取得
  const emails = sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues();
  // Calendar API用にフォーマット
  const formatedEmails = shuffle(emails).map(email => ({'email': email[0]}));
  // グループの作成
  const lunchGroups = createLunchGroups(formatedEmails);
  // 2020入社新卒カレンダーのIDをシートから取得
  const calendarId = sheet.getRange(2, 2).getValue();
  // グループごとにイベント登録
  lunchGroups.map(group => createEvent(calendarId, group));
}
