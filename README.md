# VBA 採用管理フォーム

Excel VBA で作成した採用管理フォームです。

##スクリーンショット
<img width="1130" height="599" alt="image" src="https://github.com/user-attachments/assets/eacb6792-4e51-4297-b94a-010fb8c90c4e" />

## 主な機能
- 応募者情報の登録
- カレンダーからの日付入力（動的ボタン生成）
- WithEvents を使った日付ボタンクリック処理
- コンボボックスの自動読み込み
- シートへの自動登録

## ファイル構成
- frmCalendar.frm / frx … カレンダー画面
- 採用入力フォーム.frm / frx … 応募者入力画面
- clsDayButton.cls … 日付ボタンのイベントクラス
- Module1.bas … フォーム起動用
