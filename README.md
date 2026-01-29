# generator
節能績效計畫書生成器

1.使用變數方式，取代文件中同個數值，達到數值更動時同個數值會一起變動。

2.使用excel分頁方式插入表格，若列數增加時word中表格會同步增加列數。


a.變數小數點後規則:

  "me_ ":保留原始小數

  "_rate"、 "elec_"、 "new_cop_std"、 "new_eff_std":保留2 位小數
  
  "_year":保留1 位小數

b.表格欄位名稱"kwh", "elecost", "eleccostperkwh"，才會進行數值取值至小數點後兩位，其餘皆為文字內容傳送

