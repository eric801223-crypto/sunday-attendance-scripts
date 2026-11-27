# sunday-attendance-scripts

會所主日指標牧養名單的分析與回測程式。規則文件以 claude.ai 專案「會所主日指標牧養名單」為主，`rules/` 資料夾存有副本（兩份規則文件及專案指示）；程式以本儲存庫為執行時下載的來源，專案內的 .py 為備援。

| 程式 | 用途 |
|---|---|
| sunday_attendance.py | 關心名單分析（規則 V7） |
| verify_lists.py | 關心名單回測 |
| make_cards.py | 關心名單卡片與 PDF |
| annual_attendance.py | 近一年出席分析（規則 V1.1） |
| verify_annual.py | 近一年分析回測 |
| notion_page_pdf.py | Notion 頁面 PDF（兩者共用） |
| care_html.py＋care_html_template.html | 關心名單互動版 HTML（以近12週為主） |
| annual_html.py＋annual_html_template.html | 近一年分析互動版 HTML |
| combine_html.py | 兩個互動版合成一個檔案（兩個分頁） |

| rules/主日出席分析規則_V7.md | 關心名單規則（副本） |
| rules/近一年出席分析規則_V1.md | 近一年分析規則（副本） |
| rules/專案指示.md | 專案指示（副本） |

互動版 HTML 的程式需與同名範本放在同一個資料夾。修改規則時，規則文件、程式與回測程式需同步更新本儲存庫與專案（兩份逐位元組相同），並重跑回測確認通過。
