/** 現在の翌日からn日後までの日付配列を返す */
function generateDaysChoices(n) {
  const now = new Date();
  const choices = [];
  const dayOfWeek = ['日', '月', '火', '水', '木', '金', '土'];

  // n日分の選択肢を生成
  for (let i = 1; i <= n; i++) {
    const date = new Date(now.getFullYear(), now.getMonth(), now.getDate() + i);
    const formattedDate = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy年M月d日');
    choices.push(`${formattedDate}（${dayOfWeek[date.getDay()]}）`);
  }
  return choices;
}


// /**
//  * 【トリガー】「お届け日」のプルダウン選択肢を更新
//  */
// function updateDaysChoices() {
//   try {
//     const item = form.getItems(FormApp.ItemType.LIST)[0];
//     if (item && item.getTitle() === AppConfig.form.titles.DELIVERY_DATE) {
//       item.asListItem().setChoiceValues(generateDaysChoices(30));
//     } else {
//       throw new Error("「お届け日」の日付更新に失敗しました。");
//     }
//   } catch (error) {
//     logErrorToSheet("updateDaysChoices", error);
//   }
// }