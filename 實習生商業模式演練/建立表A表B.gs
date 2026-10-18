/**
 * AAMA 實習生商業模式演練｜第 1 輪 訂閱制
 * 一鍵建立：表 A（實習生自評）、表 B（正職回饋）＋共用回應試算表
 *
 * 使用方式：
 * 1. 開啟 https://script.google.com → 新專案
 * 2. 刪掉預設內容，貼上本檔全部內容，按儲存
 * 3. 上方函式選 createForms → 執行 → 依提示授權（第一次會出現「Google 尚未驗證」，點進階 → 前往）
 * 4. 執行完成後，到「執行紀錄」看兩份表單的填答連結
 */

const ROUND = '第 1 輪｜訂閱制';
const INTERNS = ['Nuby 高瑭耘', 'Aaron 莊佾憲', 'Oscar 葉維哲', 'Wei 許巧薇'];
// 請依今天實際出席的正職調整；選項另有「其他」可自填
const STAFF = ['Stephanie', 'Fiona', 'Sharon', 'Eileen', 'Chi', 'Cheer'];
const DIMENSIONS = ['九宮格完整度與正確性', '資料佐證與研究深度', '洞察與觀點', '表達與簡報品質'];
const SCALE = ['1', '2', '3', '4', '5'];

const ANCHORS =
  '【評分錨點】\n' +
  '・九宮格完整度｜2 分：有格子沒填或放錯格／3 分：九格齊全但沒有串起來／5 分：黃金三角（客群、價值、收益）說得通，其他格都能支撐它\n' +
  '・資料佐證｜2 分：大多是臆測／3 分：有官網資料，但關鍵數字沒出處／5 分：官網與內部雲端資料交叉驗證，分清楚事實與推估\n' +
  '・洞察觀點｜2 分：只描述公司做什麼／3 分：有指出訂閱制特徵／5 分：說得出這家公司訂閱制的成功關鍵與風險\n' +
  '・表達簡報｜2 分：超時、重點不清／3 分：時間內講完但取捨不夠／5 分：5 分鐘內講到重點，QA 回應精準';

function createForms() {
  const folder = DriveApp.createFolder('實習生商業模式演練_R1訂閱制_回饋表單');
  const ss = SpreadsheetApp.create('實習生商業模式演練_R1訂閱制_回應彙整');
  DriveApp.getFileById(ss.getId()).moveTo(folder);

  const formA = buildFormA();
  const formB = buildFormB();

  [formA, formB].forEach(f => {
    f.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    DriveApp.getFileById(f.getId()).moveTo(folder);
  });

  Logger.log('資料夾：' + folder.getUrl());
  Logger.log('回應試算表：' + ss.getUrl());
  Logger.log('表 A 填答連結（實習生）：' + formA.getPublishedUrl());
  Logger.log('表 A 編輯連結：' + formA.getEditUrl());
  Logger.log('表 B 填答連結（正職）：' + formB.getPublishedUrl());
  Logger.log('表 B 編輯連結：' + formB.getEditUrl());
}

function buildFormA() {
  const f = FormApp.create('表 A｜實習生自評｜' + ROUND);
  f.setDescription(
    '請在報告結束後 5 分鐘內填寫。\n' +
    '這是你的個人成長紀錄，不對外評分；分數會和正職回饋的平均做對照，作為之後一對一回顧的參考。'
  );
  f.setProgressBar(false);
  f.setConfirmationMessage('謝謝！等一下的學習總結，請準備 2 分鐘口述。');

  f.addListItem().setTitle('姓名').setChoiceValues(INTERNS).setRequired(true);
  f.addMultipleChoiceItem().setTitle('組別').setChoiceValues(['A 組', 'B 組']).setRequired(true);

  f.addGridItem()
    .setTitle('四構面自評（1＝待加強，5＝很好）')
    .setHelpText(ANCHORS)
    .setRows(DIMENSIONS)
    .setColumns(SCALE)
    .setRequired(true);

  f.addTextItem()
    .setTitle('我負責的部分')
    .setHelpText('負責九宮格的哪幾格、簡報的哪幾頁')
    .setRequired(true);

  f.addParagraphTextItem()
    .setTitle('本輪盲點')
    .setHelpText('QA 時被問倒、或聽完同事補充後才發現自己沒想到的地方')
    .setRequired(true);

  f.addCheckboxItem()
    .setTitle('需要幫忙的地方（可複選）')
    .setChoiceValues([
      '找資料／判斷資料可不可信',
      '收益與成本怎麼推估',
      '九宮格各格之間的邏輯',
      '簡報的取捨與節奏',
      'QA 應答',
      '想請教熟悉這家校友的同仁',
    ])
    .showOtherOption(true);

  f.addParagraphTextItem()
    .setTitle('具體需要什麼幫忙？')
    .setHelpText('選填，寫越具體越好安排，例如：想請教哪位同仁、卡在哪個數字');

  return f;
}

function buildFormB() {
  const f = FormApp.create('表 B｜正職回饋｜' + ROUND);
  f.setDescription(
    '請在 00:45 填表時間內完成，約 3–4 分鐘。以「組」為單位評分。\n' +
    '回饋記名，由專案經理彙整後轉交實習生；個別分數不公開，實習生只會看到平均與文字回饋。\n' +
    '※ 打 1 分或 5 分，請在最後一題寫一個具體事例。'
  );
  f.setProgressBar(false);
  f.setConfirmationMessage('謝謝你的回饋！');

  f.addMultipleChoiceItem()
    .setTitle('評分者')
    .setChoiceValues(STAFF)
    .showOtherOption(true)
    .setRequired(true);

  ['A 組', 'B 組'].forEach(g => {
    f.addSectionHeaderItem().setTitle(g).setHelpText('評分範圍：指定家 iCHEF＋自選家的整體表現');
    f.addGridItem()
      .setTitle(g + '｜四構面評分（1–5 分）')
      .setHelpText(ANCHORS)
      .setRows(DIMENSIONS)
      .setColumns(SCALE)
      .setRequired(true);
    f.addTextItem().setTitle(g + '｜👍 最值得保留的一點').setRequired(true);
    f.addTextItem().setTitle(g + '｜🔧 最該改的一點＋建議做法').setRequired(true);
  });

  f.addSectionHeaderItem().setTitle('補充');
  f.addParagraphTextItem()
    .setTitle('💡 外部研究看不到的資訊')
    .setHelpText('熟悉該校友的同仁填寫：近況、真實收費結構、關鍵轉折。其他人可略過');
  f.addTextItem().setTitle('🙋 今天誰提出的問題最好？為什麼？');
  f.addParagraphTextItem()
    .setTitle('打 1 分或 5 分的具體事例')
    .setHelpText('有打 1 分或 5 分的話必填：請寫是哪一組、哪個構面、發生了什麼');

  return f;
}
