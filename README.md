# sunday-attendance-scripts

會所主日指標牧養名單的分析與回測程式。規則文件放在 claude.ai 專案「會所主日指標牧養名單」，本儲存庫只放程式，作為執行時下載的來源；專案內的 .py 為備援。

| 程式 | 用途 |
|---|---|
| sunday_attendance.py | 關心名單分析（規則 V7） |
| verify_lists.py | 關心名單回測 |
| make_cards.py | 關心名單卡片與 PDF |
| annual_attendance.py | 近一年出席分析（規則 V1.1） |
| verify_annual.py | 近一年分析回測 |
| notion_page_pdf.py | Notion 頁面 PDF（兩者共用） |

修改規則時，規則文件、程式與回測程式需同步更新本儲存庫與專案，並重跑回測確認通過。
