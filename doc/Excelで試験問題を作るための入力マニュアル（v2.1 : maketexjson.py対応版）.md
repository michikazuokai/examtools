# ✅ Excelで試験問題を作るための入力マニュアル（v2.1 / maketexjson.py対応版）
（更新点：**b_subquest でも b_question と同じ列（E〜H）で改ページ／行空き指定が可能**。さらに **数式入力／LaTeX直書き `[[]]`**、**maketexjson.py の引数**を追記）

---

## 0) 基本ルール（シート全体）

- **A列：タグ（必須）**  
  例：`b_exam`, `b_question`, `b_select`, `b_subquest` など
- **A列が空の行 / 行自体が空**は無視されます。
- 未対応タグは基本無視されます（ただし **使用禁止タグ**はエラーになる場合があります）。

---

## 1) 試験の表紙ブロック（b_exam ～ e_exam）

| タグ        |   列 | 内容                                                     |
| ----------- | ---: | -------------------------------------------------------- |
| `b_exam`    |    A | 表紙ブロック開始                                         |
| `examtitle` |    B | 試験タイトル                                             |
| `examnote`  |    B | 注意書き（複数行OK：行を増やして `examnote` を複数置く） |
| `subject`   |    B | 科目番号（例：2022101）                                  |
| `fsyear`    |    B | 年度（例：2025）                                         |
| `ansnote`   |    B | 解答欄の注意（任意）                                     |
| `anssize`   |    B | 解答用紙サイズ（幅,高さ）例：`(50.0,60.0)`               |
| `e_exam`    |    A | 表紙ブロック終了                                         |

---

## 2) バージョン指定（qpattern）

| タグ       |   列 | 内容                      |
| ---------- | ---: | ------------------------- |
| `qpattern` |    B | 例：`A,B`（カンマ区切り） |

- `qpattern` が **未指定 or Aのみ** → A版のみ
- `qpattern` が **A,B** → A/B 両方を生成（B版では並べ替えが有効）

---

## 3) 大問ブロック（b_question ～ e_question）【v2重要】

### 3-1) b_question 行の列仕様（v2）
`b_question` 行は **C列以降に制御列**を持ちます。

|    列 | 内容               | 説明                                        |
| ----: | ------------------ | ------------------------------------------- |
|     A | `b_question`       | 大問開始                                    |
| **C** | **qid（必須）**    | 一意ID（重複禁止）                          |
| **D** | orderB（任意）     | **B版の大問並び替えキー**（整数、小さい順） |
| **E** | PB_A_after（任意） | **A版で**この大問の後に改ページ：`1`        |
| **F** | LS_A_after（任意） | **A版で**この大問の後に行空き：数値         |
| **G** | PB_B_after（任意） | **B版で**この大問の後に改ページ：`1`        |
| **H** | LS_B_after（任意） | **B版で**この大問の後に行空き：数値         |

### 3-2) 基本タグ
| タグ         |   列 | 内容                |
| ------------ | ---: | ------------------- |
| `question`   |    B | 大問の問題文（1行） |
| `e_question` |    A | 大問終了            |

---

## 4) 大問の本文要素（content に入るもの）

### 4-1) 単独テキスト行：sline
| タグ    |   列 | 内容                                 |
| ------- | ---: | ------------------------------------ |
| `sline` |    B | 表示する1行テキスト                  |
|         |    C | (before,after) 例：`(0,0)` / `(5,0)` |

### 4-2) 複数行テキスト：b_multiline ～ e_multiline
| タグ          |   列 | 内容                       |
| ------------- | ---: | -------------------------- |
| `b_multiline` |    B | (before,after) 例：`(5,0)` |
| `text`        |    B | 本文1行（複数行OK）        |
| `e_multiline` |    A | multiline終了              |

### 4-3) 選択肢：b_select ～ e_select（B版並び替え対応）
| タグ       |   列 | 内容                                                 |
| ---------- | ---: | ---------------------------------------------------- |
| `b_select` |    B | 表示スタイル：`normal` / `inline` / `inline(8)` など |
| `select`   |    B | 選択肢本文                                           |
|            |    C | ラベル（任意）                                       |
|            |    G | order（任意）※B版で並び替えキー（整数）              |
| `e_select` |    A | 終了                                                 |

### 4-4) コードブロック：b_code ～ e_code
| タグ     |   列 | 内容                                           |
| -------- | ---: | ---------------------------------------------- |
| `b_code` |    B | `linenumber` を入れると行番号ON、それ以外はOFF |
| `code`   |    B | コード1行（複数行OK）                          |
| `e_code` |    A | 終了                                           |

### 4-5) 画像：image
| タグ    |   列 | 内容                              |
| ------- | ---: | --------------------------------- |
| `image` |    B | `ファイル名[幅]` 例：`3.png[6.5]` |

---

## 5) 小問（subquestions）：b_subgroup ～ e_subgroup 【更新点あり】

### 5-1) 小問グループの基本
| タグ         |   列 | 内容                               |
| ------------ | ---: | ---------------------------------- |
| `b_subgroup` |    A | 小問グループ開始                   |
| `b_subquest` |    A | 小問開始                           |
| `subquest`   |    B | 小問の問題文                       |
| `e_subquest` |    A | 小問終了                           |
| `e_subgroup` |    A | 小問グループ終了（大問に取り込む） |

### 5-2) ✅ b_subquest 行でも改ページ／行空き指定が可能（b_question と同じ列）
**b_subquest 行**に以下の列を入れると、その小問の直後に制御が入ります。

|    列 | 内容               | 説明                                 |
| ----: | ------------------ | ------------------------------------ |
| **E** | PB_A_after（任意） | **A版で**この小問の後に改ページ：`1` |
| **F** | LS_A_after（任意） | **A版で**この小問の後に行空き：数値  |
| **G** | PB_B_after（任意） | **B版で**この小問の後に改ページ：`1` |
| **H** | LS_B_after（任意） | **B版で**この小問の後に行空き：数値  |

> 例）A版で「小問2の後で改ページ」→ 小問2の `b_subquest` 行の **E列に `1`**  
> 例）B版で「小問3の後で1行空き」→ 小問3の `b_subquest` 行の **H列に `1`**（数値は運用に合わせて）

### 5-3) 小問内で使える要素（大問と同様）
| タグ                               |    列 | 内容                             |
| ---------------------------------- | ----: | -------------------------------- |
| `subsline`                         |   B/C | 1行テキスト（Cに(before,after)） |
| `b_submultiline`～`e_submultiline` |     B | 複数行テキスト                   |
| `b_subselect`～`e_subselect`       | B/C/G | 選択肢（大問selectと同じ）       |
| `b_subcode`～`e_subcode`           |     B | コード（大問codeと同じ）         |
| `subimage`                         |     B | 画像（`ファイル名[幅]`）         |

---

## 6) 改ページ・行空きの指定方法（v2）

### 6-1) 改ページは「after列」で指定（推奨運用）
- **大問の後**：`b_question` 行の **E〜H**
- **小問の後**：`b_subquest` 行の **E〜H**

### 6-2) どこに `1` を入れるか（重要）
- 「次ページの先頭にしたい問題」の直前、つまり  
  **“区切りの最後の問題” の PB_*_after を `1`** にする。

---

## 7) LINESPACE タグ（制約）
- `LINESPACE` は **subgroup 内でのみ使用可能**です。
- subgroup 外で `LINESPACE` を使うとエラーになります。

| タグ        |   列 | 内容                                 |
| ----------- | ---: | ------------------------------------ |
| `LINESPACE` |    B | 行数（数字）例：`2`（未指定は1扱い） |

---

## 8) 最低限テンプレ（例）

1. `b_exam` → `examtitle/subject/fsyear/...` → `e_exam`
2. `qpattern`（例：A,B）
3. `b_question`（C列qid、必要ならE〜HでPB/LS）  
   → `question`  
   → 本文（`sline` / `b_multiline` / `b_select` / `b_code` / `image` など）  
   → 小問があるなら `b_subgroup` … `e_subgroup`  
      - 小問 `b_subquest` の **E〜Hで改ページ/行空きが可能**  
   → `e_question`
4. 次の大問へ

---

## 9) LaTeXの数式の記入方法（Excel入力）

Excelのセルに **LaTeXの数式**を書いて、PDF（LaTeX）で正しく表示させるためのルールです。

### 9-1) インライン数式（文章の途中）
- `\( ... \)` または `$ ... $` を使います。  
  例：`平均は \( \mu \) とする`、`平均は $ \mu $ とする`

### 9-2) 別行立て数式（大きく中央に表示）
- `\[ ... \]` を使います。  
  例：`\[ \bar{x} = \frac{1}{n}\sum_{i=1}^{n} x_i \]`

### 9-3) displaystyle を使いたい場合
- 例：`\( \displaystyle \frac{3}{27} \)`

> 注意：通常入力のまま `\(` などを書くと、変換処理の影響で崩れる可能性があるため、次の `[[]]` 方式を推奨します。

---

## 10) LaTeXを直記入できる `[[]]` の書き方（重要）

通常、Excelの文字列は LaTeX生成時に **エスケープ処理**（`\` や `{` などの保護）が入ります。  
数式やLaTeX命令をそのまま書きたい場合は **`[[ ... ]]` で囲む**と、中身を **生のLaTeXとしてそのまま出力**します。

### 10-1) 基本ルール
- `[[ ... ]]` の中は **エスケープされません**
- `[[ ... ]]` の外側は **通常どおりエスケープされます**
- 文章中に混在できます

### 10-2) 例：インライン数式を安全に書く
- `平均は [[\( \mu \)]] とする`
- `確率は [[\( P(A) \)]] で表す`

### 10-3) 例：分数（displaystyle）を確実に大きく出す
- `標準誤差は [[\( \displaystyle \frac{\sigma}{\sqrt{n}} \)]] である`

### 10-4) 例：別行立て数式（中央に大きく表示）
- `[[\[ \bar{x} = \frac{1}{n}\sum_{i=1}^{n} x_i \]]]`

### 10-5) 例：LaTeX命令を直書き（強調・下線など）
- `この式は [[\underline{重要}]] である`
- `[[\textbf{注意}]]：選択問題は1つだけ選ぶ`

### 10-6) 注意事項（事故防止）
- `[[ ... ]]` 内の `%` はコメント扱いになりやすい（意図せず消える）
- `{}` の閉じ忘れはコンパイルエラーの原因
- `\begin{...}` / `\end{...}` を書く場合は環境整合性に注意（途中に入れると崩れる）

---

## 11) 運用ルール（統一方針）：数式とLaTeX直書きの使い分け（推奨）

この教材では、Excel → JSON → LaTeX 変換の途中で **`\` や `{}` がエスケープ処理に巻き込まれて崩れる事故**を避けるため、**数式・LaTeX命令は `[[]]` に統一**します。

### 11-1) 統一ルール（結論）
✅ **原則：`\` を含むものはすべて `[[]]` に入れる**

| 入力したい内容               | 推奨する書き方（Excelセル）                       | 備考              |
| ---------------------------- | ------------------------------------------------- | ----------------- |
| 文章だけ                     | `平均は135である。`                               | 通常入力でOK      |
| インライン数式               | `平均は [[\( \mu \)]] とする`                     | **必ず `[[]]`**   |
| 分数を大きく（displaystyle） | `[[\( \displaystyle \frac{3}{27} \)]]`            | 事故防止で `[[]]` |
| 別行立て数式                 | `[[\[ z=\frac{\bar{x}-\mu}{\sigma/\sqrt{n}} \]]]` | **必ず `[[]]`**   |
| 太字・下線など命令           | `[[\textbf{注意}]]：...` / `[[\underline{重要}]]` | **必ず `[[]]`**   |
| LaTeXの記号（\%, \_, \# 等） | `[[\%]]` / `[[\_]]` / `[[\#]]`                    | `%` は特に注意    |

### 11-2) 書き方テンプレ（コピペ用）
- インライン数式：`[[\( ... \)]]`
- 別行数式：`[[\[ ... \]]]`
- displaystyle：`[[\( \displaystyle ... \)]]`

### 11-3) 禁止・非推奨（事故が起きやすい）
- 🚫 非推奨：通常入力のまま `\( ... \)` を書く  
  例：`平均は \( \mu \) とする`
- 🚫 `%` を `[[]]` 外で使う（コメント扱いで欠落が起きる）



## 12) maketexjson.py 実行時の引数（使い方）

maketexjson.py は、Excel（試験問題.xlsm など）から **v2形式の JSON** を生成するスクリプトです。  
実行時に指定できる引数は **2つ**です。

---

### 12-1) コマンド形式

python maketexjson.py <sheetname> [excel_filename]

---

### 12-2) 引数の説明

#### ① <sheetname>（必須）
- Excelブック内の **シート名** を指定します。
- 例：2022101

例：
python maketexjson.py 2022101

#### ② [excel_filename]（任意）
- 読み込むExcelファイル名（拡張子まで）を指定します。
- 省略すると **試験問題.xlsm** を読み込みます。

例：
python maketexjson.py 2022101 試験問題_2025.xlsm

---

### 12-3) 入力ファイルの場所（重要）
スクリプトは、指定した Excel を次の場所から探します。

- ../input/<excel_filename>

例：
- ../input/試験問題.xlsm
- ../input/試験問題_2025.xlsm

---

### 12-4) 出力ファイル
生成される JSON は次に出力されます。

- ../work/<sheetname>.json

例：
- ../work/2022101.json

---

### 12-5) 生成されるバージョン（A/Bなど）
Excel内の qpattern に応じて、JSONが single / multi になります。

- qpattern が無い、または A のみ → versionmode: "single"（Aのみ）
- qpattern が A,B のように複数 → versionmode: "multi"（A/Bを生成）

※ multi の場合は、coverタイトルに (A) / (B) が自動付与されます。

---

### 12-6) 終了コード（エラーの見方）
- 正常終了：0
- 入力不備（contract validation など）：2（src付きでエラー表示されます）



## 13) makelatex.py 実行時の引数（使い方）

makelatex.py は、`work/<sheetname>.json`（maketexjson.py で生成したJSON）から **LaTeX（body.tex）** を生成するスクリプトです。  
A/B など複数バージョンがある場合は、**引数指定なしで全バージョンを自動生成**できます。:contentReference[oaicite:0]{index=0}

---

### ✅ コマンド形式

python makelatex.py <sheetname> [options]

---

## 1) 必須引数

### ① `<sheetname>`（必須）
- `work/<sheetname>.json` を読み込みます（既定）。
- 例：`2022101`

例：
python makelatex.py 2022101

---

## 2) オプション引数

### ② `--version <A|B|...>`（任意）
- 出力するバージョンを指定します。
- **省略した場合：JSON内の versions から自動取得し、全バージョンを生成**します。

例（B版だけ生成）：
python makelatex.py 2022101 --version B

例（A版だけ生成）：
python makelatex.py 2022101 --version A

---

### ③ `--in <input_json_path>`（任意）
- 入力JSONファイルのパスを指定します。
- 省略時：`../work/<sheetname>.json`

例：
python makelatex.py 2022101 --in ../work/2022101.json

---

### ④ `--out <output_tex_path>`（任意）
- **出力先を直接指定**します。
- 注意：このオプションはあるが、現行コードでは通常ルートの出力を上書きする運用になりやすいので、
  基本は使わず、既定の出力先（後述）を推奨します。

例：
python makelatex.py 2022101 --version A --out ./tmp/body.tex

---

### ⑤ `--cover`（任意）
- JSON内に cover 要素があれば、それを LaTeX に出力する（※実装は「coverコメント」と説明されています）。
- 通常は **texmerge/コンパイル側で表紙を付ける**運用のため、不要なら付けません。

例：
python makelatex.py 2022101 --version A --cover

---

### ⑥ `--no-trace`（任意）
- Excel行番号などの追跡用コメント（`%% QBEGIN ...` 等）を出力しません。
- 通常デバッグ目的では **traceあり推奨**。

例：
python makelatex.py 2022101 --version A --no-trace

---

## 3) 入力ファイルの場所（既定）
- 既定入力：`../work/<sheetname>.json`

例：
- `../work/2022101.json`

---

## 4) 出力ファイルの場所（既定）

バージョンごとに次の場所へ出力します：

- `../output/<sheetname>/<version>/<sheetname>_<version>_body.tex`

例：
- `../output/2022101/A/2022101_A_body.tex`
- `../output/2022101/B/2022101_B_body.tex`

---

## 5) よく使う実行例（まとめ）

### ✅ 例1：全バージョンを生成（A/Bがあれば両方）
python makelatex.py 2022101

### ✅ 例2：B版だけ生成
python makelatex.py 2022101 --version B

### ✅ 例3：入力JSONを明示して生成
python makelatex.py 2022101 --in ../work/2022101.json

### ✅ 例4：トレースなしで生成
python makelatex.py 2022101 --version A --no-trace



## 14) texmerge_compile.py 実行時の引数（使い方）

texmerge_compile.py は、`body.tex`（makelatex.py が生成）とテンプレート（templates/latex）を結合して
**コンパイル可能な .tex を作成し、lualatex で PDF まで生成**するスクリプトです。:contentReference[oaicite:0]{index=0}

---

### ✅ コマンド形式

python texmerge_compile.py <sheet> [options]

---

## 1) 必須引数

### ① <sheet>（必須）
- 科目ID（シート名）を指定します。
- 例：`2022101`

例：
python texmerge_compile.py 2022101

---

## 2) オプション引数

### ② --version <A|B|...>（任意）
- 指定した **1つの版だけ**をコンパイルします。
- 省略した場合は `work/<sheet>.json` を読み、`versions[].version` から版（A/B…）を取得して **全版を処理**します。

例（B版だけ）：
python texmerge_compile.py 2022101 --version B

---

### ③ --runs <N>（任意）
- lualatex を回す回数を指定します（レイアウト確定のため **2推奨**）。
- 省略時の既定値：`2`

例（2回回す）：
python texmerge_compile.py 2022101 --runs 2

---

## 3) 入力の探索ルール（自動）

### 3-1) 版の一覧（A/Bなど）
- `work/<sheet>.json` を読み、`versions[].version` から取得します。
- `versions` が無い場合は `["A"]` とみなします。

---

### 3-2) body.tex の探索順
次の候補を上から順に探し、見つかったものを使います。

1. `output/<sheet>/<ver>/<sheet>_<ver>_body.tex`
2. `work/<sheet>_<ver>_body.tex`
3. `output/<sheet>/<ver>/body.tex`

---

## 4) 出力（生成されるファイル）

版ごとに次のディレクトリへ出力します。

- `output/<sheet>/<ver>/`

生成物の例：
- `output/2022101/B/2022101_B_body.tex`（bodyをコピーしたもの）
- `output/2022101/B/2022101_B.tex`（テンプレ＋bodyの結合後tex）
- `output/2022101/B/2022101_B.pdf`（コンパイル結果）

---

## 5) 画像の参照（graphicspath 自動注入）

- `images/<sheet>/` を参照するように `preamble.tex` に自動注入します。
- これにより、本文で `\includegraphics{5.png}` のように **ファイル名だけ**で参照できます。

---

## 6) よく使う実行例（まとめ）

### ✅ 例1：全版（A/B）をまとめてPDF化（推奨：runs=2）
python texmerge_compile.py 2022101 --runs 2

### ✅ 例2：B版だけPDF化
python texmerge_compile.py 2022101 --version B --runs 2

---

## 15) anstest_ans.py 実行時の引数（使い方）

anstest_ans.py は、Excel（既定：`input/試験問題.xlsm`）から **解答用JSON** を生成し、
さらに **解答用紙PDF** を出力するスクリプトです。:contentReference[oaicite:0]{index=0}

---

### ✅ コマンド形式

python anstest_ans.py <subject> [options]

---

## 1) 必須引数

### ① <subject>（必須）
- 科目番号（= Excelの **シート名**）を指定します。
- 例：`2022101`

例：
python anstest_ans.py 2022101

---

## 2) オプション引数

### ② --version <A|B|A,B>（任意）
- 出力するバージョンを指定します（`A` / `B` / `A,B`）。
- 省略した場合は、Excel内の `qpattern` に従い **利用可能な全バージョンをPDF出力**します。

例（B版だけPDF出力）：
python anstest_ans.py 2022101 --version B

例（A,BをPDF出力）：
python anstest_ans.py 2022101 --version A,B

---

## 3) 入力ファイル（固定）
このスクリプトは、Excelファイル名を引数で受け取らず、次を固定で読みます。

- 入力Excel：../input/試験問題.xlsm

（例：scripts/ 配下から実行している場合、プロジェクトの `input/試験問題.xlsm`）

---

## 4) 出力ファイル

### 4-1) 解答JSON
- 出力先：../work/answers_<subject>.json

例：
- ../work/answers_2022101.json

※ JSONは **常に全バージョン分（single/multi）を保持**する形式で出力されます。  
（--version は主にPDF出力の対象を絞るための指定です）

### 4-2) 解答用紙PDF
- 出力先：../output/<subject>/<version>/<subject>_<version>_解答用紙.pdf

例：
- ../output/2022101/A/2022101_A_解答用紙.pdf
- ../output/2022101/B/2022101_B_解答用紙.pdf

---

## 5) 生成されるバージョン（single / multi）
Excel内の `qpattern` に応じて、解答JSONは次になります。

- qpattern が無い / Aのみ → versionmode: "single"（Aのみ）
- qpattern が A,B など複数 → versionmode: "multi"（A/Bなど）

---

## 6) よく使う実行例（まとめ）

### ✅ 例1：全バージョンを出力（qpatternに従う）
python anstest_ans.py 2022101

### ✅ 例2：B版だけ解答用紙PDFを出力
python anstest_ans.py 2022101 --version B