# 交流会進捗管理の生成元（tools）

庁内へ持ち込む `30_deliverables/package_src/` の2本（`build_progress.bat`、`build_progress.ps1`）と、A4 1枚の手順書を作る。
設計の正は `../20_plans/261006_ps1で組み立てる計画.md`。手本は `../../R8静岡市交流会/tools/260924_照合用Excel/`。
**直すのはここ。** `package_src` の ps1 を直接直さない（作り直すと上書きされる）。

## 作り直しかた

```
cd tools
python build_ps1.py
python build_tejun.py
node %USERPROFILE%\.claude\skills\ponchi\build.mjs ..\30_deliverables\261006_手順書_交流会進捗管理を業務端末で作り直す.html --pdf
powershell -ExecutionPolicy Bypass -NoProfile -File "%USERPROFILE%\.claude\tools\make-package.ps1" -Source ..\30_deliverables\package_src -Name progress -OutDir ..\30_deliverables\_package
```

`_package` に古い zip が残っていたら、渡す前に `90_archive` へ移す（最新だけを置く）。

## どこに何があるか

| もの | 何か |
|---|---|
| `build_ps1.py` | 庁内用の ps1 と bat を組み立てる。ps1 の中身（シート、表、式、ボタン、使い方の文）はここに書いてある |
| `vba/進捗管理.vba` | 標準モジュール（UTF-8。庁内で ps1 が cp932 の `.bas` に書き出す） |
| `vba/clsBookEvents.vba` | クラスモジュール。プルダウンの切り替え、保存の直前、シートから離れたとき、を受ける |
| `vba/frmAddTask.vba` | 入力フォームのコード。部品は開くときにコードで並べる |
| `vba/frmAddTask.frx.b64` | 空のフォームの `.frx`（2026-10-06 に私用PCで VBComponents.Add(3) → Export して作った）。ps1 に埋め込み、庁内で書き戻す |
| `build_tejun.py` | 手順書の HTML。色と書体は照合用Excelの `lib_brand.py` を借りている |
| `試験/` | 私用PCの試験。庁内へは持ち込まない |
| `VBA/`、`build/` | 2026年7〜8月のもの。いまの組み立てでは使っていない（`build/` の ps1 は私用PCのパス入りで、庁内では動かない） |

`build_ps1.py` は作る前に止まる検査を持っている。モジュールの宣言が手続きより後ろ、cp932 で化ける字（〜など）、
行頭が `'@` の行（ps1 のヒアストリングが切れる）、貼り付け用に `Attribute` や `VERSION` の行が残っていること。

## ps1 が庁内でやること

1. 同じフォルダの `★交流会進捗管理★.xlsm` を、**zip の中の XML として読む**（Excel で開かない）。★は書き換えない。
   Excel で開くと、改ページプレビューで保存された★は、既定プリンターが応答しない端末で止まるため
2. 表 ProgressTable の8列を**列の名前で**写す。見つからない列があれば止まり、`作成結果.txt` に違いを書く。
   CategoryMaster、StatusMaster、担当者リストも写す。時期が「21日」のような行は「開催日」に読み替え、結果に書く
3. マクロを `マクロ` フォルダに書き出す（取り込み用3つ、貼り付け用3つ、`.frx`）
4. 新しいブックを組み立て、写した中身を読み戻して1つずつ照らし合わせる
5. `VBProject.Import` を試す（通らなくても止めない。結果に手で入れる手順を書く）。ボタンはどちらでも置く
6. 全シートを標準の表示にして `交流会進捗管理.xlsm` で保存。すでにあれば上書きせずに止まる

## 試験（私用PC）

★の実データが入るので、作業フォルダは案件の外（scratchpad）にする。

```
powershell -File 試験\試験_組み立て.ps1 -Star ..\20_plans\★交流会進捗管理★.xlsm -Out <作業フォルダ> -Route 1   (2、3 も)
python 試験\照合_写しと★.py <作業フォルダ>\★交流会進捗管理★.xlsm <作業フォルダ>\交流会進捗管理.xlsm
powershell -File 試験\試験_動き.ps1 -Book <作業フォルダ>\交流会進捗管理.xlsm
```

`照合_写しと★.py` は ps1 とは別の道具（openpyxl）で★と比べる。ps1 自身の照らし合わせは、自分が読んだものとしか比べないので、
読み違いを見つけられない（2026-10-06、セルの中の改行が CR LF のまま写って `_x000D_` になったのを、こちらで見つけた）。

## 分かったこと（2026-10-06）

- **保存の最中は、シートを切り替えられない。**BeforeSave の中で `Activate` しても黙って効かず、エラーも出ない。
  そのため保存の直前に標準の表示へ戻せるのは、開いているシートだけ。ほかのシートは、離れるとき（SheetDeactivate）に戻す
- ★の XML は、セルの中の改行が CR LF のまま書かれている。`XmlDocument.LoadXml` は XML の決まりどおりに LF へそろえないので、読む前にそろえる
- 印刷用の行の高さは、自動の高さ（AutoFit）のままだと紙で下が切れる。2割足している
