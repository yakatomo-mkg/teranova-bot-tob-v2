/** 現在(関数が実行された時点)の翌日の日付からn日後までの日付を要素に持つ配列を返す関数 */
function generateDaysChoices(n) {
  const now = new Date();
  const choices = [];
  const dayOfWeek = ['日', '月', '火', '水', '木', '金', '土'];

  // 30日分の選択肢を生成
  for (let i = 1; i <= n; i++) {
    const date = new Date(now.getFullYear(), now.getMonth(), now.getDate() + i);
    const formattedDate = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy年M月d日（' + dayOfWeek[date.getDay()] + '）');
    choices.push(formattedDate);
  }

  return choices;
}

/** 
 * 【トリガーを設定中】
 * 日付トリガーで毎日呼び出されて、「お届け日」　のプルダウン選択肢を更新する関数
 */
function updateDaysChoices() {
  try {
    const item = form.getItems(FormApp.ItemType.LIST)[0];
    console.log(item.getTitle());

    if (item && item.getTitle() === 'お届け日') {
      item.asListItem().setChoiceValues(generateDaysChoices(30));
    } else {
      throw new Error("「お届け日」の日付更新に失敗しました。")
    }
  } catch (error) {
    logErrorToSheet("updateDaysChoices", error);
    // showErrorDialog("updateDaysChoices", error);
  }
}
