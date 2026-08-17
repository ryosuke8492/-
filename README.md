# 書類照合ツール（文字抽出 基本モジュール）

事務書類（PDF）から特定項目の値を抽出し、書類間で突き合わせて照合するための
基本ロジック集です。GUI版（`doc-check-tool.html` / PDF.js）と同じ考え方を、
VBA（Excel）版・Node.js版でも使えるようにしています。

## 構成

```
vba/PDF文字抽出_基本.bas   … Excel VBA版。Wordを使ってPDFを全文テキスト化 →
                              正規表現でキーワード近くの値を抽出 → 判定
js/pdfExtractor.js         … Node.js版。PDF.jsで座標範囲を指定して値を抽出 →
                              判定（GUI版と同じロジックをNode向けに切り出したもの）
js/package.json
```

## 使い方（VBA版）

1. Excelで `Alt+F11` → VBEを開く
2. `ファイル` → `ファイルのインポート` で `PDF文字抽出_基本.bas` を取り込む
3. シートのA列にPDFのフルパス、B列に抽出したい項目の正規表現パターンを入力
   （例: `数量[：:]\s*([0-9,]+)`）
4. `Run_ExtractSample` マクロを実行 → C列に抽出結果が入る
5. `JudgeValues(値1, 値2, "exact")` のように呼び出せば、2つの値の一致判定もできる
6. 事前準備はPCにMicrosoft Wordが入っていればOK（PDFのテキスト変換にWordを利用）

## 使い方（Node.js版）

```bash
cd js
npm install
node pdfExtractor.js sample.pdf   # 全文抽出のサンプル実行
```

`extractTextInBox(pdfPath, pageNum, {x,y,w,h})` で座標指定抽出、
`judgeValues(a, b, type)` で一致判定ができます。GUI版のフィールド定義
（doc-check-tool.htmlで枠取りしたx/y/w/h）をそのまま渡せます。

## 汎用化の方向性（次のステップ）

- **Excel化**: VBA版をベースに、「書類パス・項目名・照合タイプ・期待値」を
  一覧管理する設定シートを作り、ボタン一つで全件チェック→結果シート出力、
  の形にすると汎用ツールとして使い回しやすくなります。
- **デスクトップアプリ化**: Node.js版（pdfExtractor.js）をコアロジックとして
  Electronでラップし、`doc-check-tool.html` のGUI（PDFプレビュー・枠取り・
  照合グループ・エビデンス確認）をそのまま画面として載せる。ブラウザ版と違い
  ローカルファイルへのアクセスや、複数PDFの一括処理がしやすくなります。

## 関連

- GUI版プロトタイプ: `doc-check-tool.html`（PDF.js使用、ブラウザ単体で動作）
