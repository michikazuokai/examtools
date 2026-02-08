了解。**いまのマクロ体系で「行間（縦方向スペース）」に関わる場所**を、**どこがどこに効くか／どこの数値を変えるとどう変わるか**で一覧にします。
（※あなたがここまで触ってきた `question / subquestion / choices(inline/normal) / multiline` を中心に整理）

------

## ✅ 0. 基本ルール（積み上がる理由）

- **縦方向の空きは「前のブロックの末尾」＋「次のブロックの先頭」で足される**ことがある
- `\vspace*{...}` はそのまま足されやすい
- `\addvspace{...}` は **直前の skip と比較して重複しにくい**（積み上がり防止）

------

# ✅ 行間設定一覧（どこを変えると何が変わる？）

## A) question（大問・小問の設問文ブロック）

### A-1. question の前（前の問題→次の問題）

- **関与箇所**：`question` 環境 begin 冒頭の `\addvspace{...}`
- **変える値**：
  - `\QBefore` … 大問の前の空き
  - `\QBeforeSub` … subquestion 内の「2問目以降」の前の空き（あなたが欲しかったやつ）
- **変えるとどうなる？**
  - `\QBefore` を増やす → **大問どうしの間が広がる**
  - `\QBeforeSub` を増やす → **subquestion 内の小問どうしの間が広がる**
  - 「1問目は空けない」「改ページ直後は空けない」は **フラグ判定側で制御**

### A-2. question の後（設問文→次のブロック）

- **関与箇所**：`question` 環境の本文出力直後（end 直前）の `\addvspace{\QAfter}` がある場合
- **変える値**：
  - `\QAfter`
- **変えるとどうなる？**
  - `\QAfter` を増やす → **設問文の直後に空きが出る**
  - ただし、次が `choices` なら `choices` 側の `examChoiceVspace` も足され得るので注意（足し算になる）

------

## B) choices（citem 群：normal / inline 共通）

### B-1. choices の前（直前ブロック→選択肢の先頭）

- **関与箇所**：`choices` 環境 begin の
  - `\par\addvspace{\examChoiceVspace}`
- **変える値**：
  - `examChoiceVspace`（keys の `vspace=`。デフォルト `3mm`）
- **変えるとどうなる？**
  - ここを増やす → **設問文や code の直後〜選択肢の先頭までが広がる**
  - ここを 0 にする → **設問文や code の直後に選択肢が詰まる**
- **おすすめ運用**
  - 「境界の空き」はなるべく **ここ1か所に集約**（`-\topskip` 等の“引っ張り”は使わない）

### B-2. normal choices の citem 間（縦の項目間）

- **関与箇所**：normal の `\begin{itemize}[...]`
- **変える値**：
  - `itemsep=...` … **citem と citem の間**
  - `topsep=...` … リストの最初/最後の余白（上下）
  - `parsep, partopsep` … 段落境界の追加余白（基本 0 推奨）
- **変えるとどうなる？**
  - `itemsep` を増やす → **選択肢の行間が広がる**
  - `topsep` を増やす → **選択肢ブロックの上下が広がる**

### B-3. inline choices の “citem 間”（横並びの間隔）

- **関与箇所**：inline の `itemize*` オプション
- **変える値**：
  - `itemjoin` / `itemjoin*` の `\hspace*{\examChoiceSep mm}`
  - `examChoiceSep`（keys の `sep=`。デフォルト 8）
- **変えるとどうなる？**
  - `sep` を増やす → **横並びの選択肢の間隔が広がる**
  - `sep` を減らす → **詰まる**

### B-4. inline choices の “重なり” に関与するもの（重要）

- **関与箇所**：
  - `\par\addvspace{0pt}`（段落確定）… **重なり防止に効く**
  - `\vspace*{-\topskip}`（上に引っ張る）… **重なりの原因になりやすい**
- **運用**
  - 重なりを避けたいなら **`-\topskip` は使わず**、`examChoiceVspace` で整える

------

## C) subquestion（小問題ブロックの枠）

### C-1. 親 question → subquestion の間

- **関与箇所**：`subquestion` begin の
  - `\par\addvspace{\SubBlockBefore}`
- **変える値**：
  - `\SubBlockBefore`
- **変えるとどうなる？**
  - 親の大問文から、小問ブロックに入る前の空きが変わる

### C-2. subquestion 終了 → 次の大問（次の question）まで

- **関与箇所**
  - `subquestion` end の `\par\addvspace{\SubBlockAfter}`
  - 次の `question` begin の `\QBefore`
- **変える値**
  - `\SubBlockAfter` と `\QBefore` の**両方**
- **変えるとどうなる？**
  - ここは足し算になりやすいので、どちらかに寄せて調整するのが管理しやすい

### C-3. subquestion 内の「1つ目の小問は空けない／2つ目以降は空ける」

- **関与箇所**
  - `question` begin の分岐（`ifInSubquestion` + `ifFirstSubQ`）
- **変える値**
  - `\QBeforeSub`（2つ目以降）
- **変えるとどうなる？**
  - subquestion 内の小問の間隔だけを独立にコントロールできる

------

## D) multiline（mline の複数行ブロック）

### D-1. multiline 内の行間（mline と mline の間）

- **関与箇所**：`multiline` 環境／`\mline` 定義側（マクロによる）
- **変える値**：`baselineskip`／`\jot`／環境内 `\setlength{\baselineskip}{...}` など（実装依存）
- **変えるとどうなる？**
  - multiline の「行と行」の間が変わる

### D-2. multiline → choices の境界（重なり・詰まり対策）

- **関与箇所**：`choices` begin の `\ifAfterMultiline ... \vspace*{...}`
- **変える値**
  - `\MLtoChoiceTightNormal`（normal）
  - `\MLtoChoiceTightInline`（inline）
- **変えるとどうなる？**
  - ここが **マイナス方向**だと「詰まり」や「重なり」の原因になる
  - まず 0 にして安全化 → その後必要なら微調整

------

## E) code（lstlisting/minted）→ choices

- **関与箇所**
  - code 環境の下余白（listings の `belowskip` など）
  - - `choices` の `examChoiceVspace`
- **変える値**
  - `\lstset{belowskip=...}`（listings）
  - `vspace=`（choices）
- **変えるとどうなる？**
  - 両方が足されるので「広すぎる」場合は片方に寄せて調整

------

# ✅ 実務でのおすすめ（混乱しない調整ルール）

- **境界の空き（ブロック間）**は原則「次のブロックの先頭」側に寄せる
  → `question` 後より `choices` 前の `examChoiceVspace` で調整する、など
- **マイナスの vspace（引っ張り）**は最後の手段（重なりやすい）
- `\vspace` より `\addvspace` を優先（積み上がり防止）

------

必要なら、この一覧を「**あなたの `macros.tex` の該当行番号つき**」にして、
「ここ（行xx）を変えるとここが変わる」と **完全にピンポイント化**した版も作れます。
その場合は、いま使っている `macros.tex` をこちらで参照して行番号を付け直します。