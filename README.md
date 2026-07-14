# docxpdf-native

**DOCXを、Pythonだけで、そのままPDFに。**

`docxpdf-native`は、DOCXの中身（Office Open XML）を直接読み取り、独自にページ配置を計算してPDFを生成するPythonライブラリです。Microsoft WordもLibreOfficeもブラウザーも外部レンダラーもクラウドAPIも呼び出しません。変換中のネットワーク接続も一切不要です。

「Wordをサーバーに入れられない」「変換のたびにネットへ出したくない」「同じ入力からいつも同じPDFがほしい」——そんな場面のために作られています。

ひとつだけ、最初にお伝えしたい大切なことがあります。このライブラリは**仕様を限定したDOCXのサブセット**を対象としており、Microsoft Wordとの完全互換は保証しません。対応範囲外の機能に出会ったときは、strictモードなら正直に停止し、lenientモードなら警告として記録します。「何ができて、何ができないか」を隠さないことが、このライブラリの設計方針です。

## 目次

- [特徴](#特徴)
- [動作要件](#動作要件)
- [インストール](#インストール)
- [はじめての変換](#はじめての変換)
- [日本語文書を変換する](#日本語文書を変換する)
- [CLIリファレンス](#cliリファレンス)
- [Python API](#python-api)
- [strictモードとlenientモード](#strictモードとlenientモード)
- [困ったときは](#困ったときは)
- [対応している機能](#対応している機能)
- [未対応の機能](#未対応の機能)
- [既知の制限](#既知の制限)
- [ページ数と決定性](#ページ数と決定性)
- [レイアウトJSONと診断JSON](#レイアウトjsonと診断json)
- [セキュリティ上の配慮](#セキュリティ上の配慮)
- [開発に参加する](#開発に参加する)
- [公式PDF回帰ベンチマーク](#公式pdf回帰ベンチマーク)
- [ライセンス](#ライセンス)

## 特徴

- **完全ネイティブ変換** — WordやLibreOfficeのインストール、外部プロセスの起動、ネットワーク接続はすべて不要です。
- **決定的な出力** — 同じ入力・設定・フォント・依存バージョンなら、レイアウトJSONもPDFバイト列も再現されます。
- **正直な診断** — 未対応機能やフォント置換は、位置と回避策つきで記録されます。「なんとなく変になった」を残しません。
- **日本語への配慮** — 行頭・行末禁則、NFC正規化、East Asiaフォント解決、実フォントメトリクスによる計測を行います。
- **信頼できない入力を前提とした設計** — ZIP爆弾対策、XML外部実体の拒否、リソース上限など、公開サービスへの組み込みを意識しています。

実行時の直接依存は3個だけです。

| パッケージ | 用途 |
| --- | --- |
| `pydantic` | 不変の文書・レイアウト・診断モデル |
| `reportlab` | 計算済みレイアウトの低レベルPDF描画 |
| `fonttools` | フォント名の取得と実フォントメトリクスの計測 |

ReportLabのPlatypusは使いません。行分割、ページ分割、表配置、フォント置換はすべてこのライブラリ自身が決定します。

## 動作要件

- Python 3.13以上
- パッケージ管理と実行には [`uv`](https://docs.astral.sh/uv/) を使用します

## インストール

リポジトリのルートで次の2行を実行するだけで準備完了です。

```console
uv sync
uv run docxpdf-native --version
```

バージョン番号が表示されれば、インストールは成功しています。

別の`uv`プロジェクトからローカルチェックアウトを使いたい場合は、パスを指定して追加してください。

```console
uv add --editable ../docxpdf-native
```

なお、このリポジトリでは`requirements.txt`は使いません。依存とツール設定は`pyproject.toml`に、解決結果は`uv.lock`にあります。

## はじめての変換

まずはCLIで1ファイル変換してみましょう。`--output`（短縮形`-o`）は必須です。

```console
uv run docxpdf-native input.docx --output output.pdf
```

成功すると、変換結果とページ数が表示されます。

```console
Converted input.docx -> output.pdf (3 pages)
```

Pythonからはこう書きます。

```python
from pathlib import Path

from docxpdf_native import ConversionOptions, Converter

converter = Converter(options=ConversionOptions())
result = converter.convert(
    source=Path("input.docx"),
    destination=Path("output.pdf"),
)

print(result.page_count)          # 生成されたページ数
print(result.warnings)            # 変換中の警告
print(result.font_substitutions)  # 行われたフォント置換
```

もしここでエラーになっても大丈夫です。[困ったときは](#困ったときは)に終了コードごとの対処をまとめてあります。特に日本語文書では、次のセクションのフォント設定が必要になることが多いです。

## 日本語文書を変換する

日本語のDOCXをきれいに変換する鍵は、**DOCXが要求するフォントと同じ（または対応する）フォントファイルを用意すること**です。フォントが変わると字幅が変わり、行数、改ページ、最終的なページ数まで変わってしまいます。

たとえば「MS 明朝」「MS ゴシック」を使った文書を、Notoフォントで変換する例です。

```python
from pathlib import Path

from docxpdf_native import ConversionOptions, Converter, FontConfiguration

fonts = FontConfiguration(
    registered_fonts={
        "Noto Sans CJK JP": Path("fonts/NotoSansCJKjp-Regular.ttf"),
    },
    font_directories=(Path("fonts"),),
    substitutions={
        "MS Mincho": "Noto Serif CJK JP",
        "MS Gothic": "Noto Sans CJK JP",
    },
    default_font="Noto Sans CJK JP",
)

options = ConversionOptions(
    strict=True,
    deterministic=True,
    font_configuration=fonts,
)
result = Converter(options=options).convert("input.docx", "output.pdf")
```

CLIでも同じことができます。

```console
uv run docxpdf-native input.docx \
  --output output.pdf \
  --font-dir ./fonts \
  --font-substitution "MS Mincho=Noto Serif CJK JP" \
  --font-substitution "MS Gothic=Noto Sans CJK JP"
```

知っておくと役に立つポイントをいくつか。

- **フォントの検索順**は、APIで登録したフォント（`registered_fonts`）、`font_directories`、環境変数、OS標準フォントディレクトリ、ReportLab標準フォント、明示的な置換、既定フォントの順です。
- 環境変数の既定名は`DOCXPDF_NATIVE_FONT_DIRS`で、複数のパスはOSのパス区切り文字（macOS/Linuxは`:`）で区切ります。
- 特定のフォントファイルに別名を付ける`registered_fonts`はPython API専用です。CLIの`--font-dir`はディレクトリ検索の追加だけを行います。
- 環境間で結果を完全に一致させたい場合は、`include_system_fonts=False`にしてOSのフォントに依存しない構成にするのがおすすめです（`examples/custom_fonts.py`がこの構成です）。
- lenientモードの既定フォールバックであるHelveticaは**日本語グリフを持ちません**。日本語文書では必ず明示的な置換先か`default_font`を用意してください。

`examples/`にそのまま実行できる例を用意しています。

```console
uv run python examples/basic_conversion.py input.docx output.pdf
uv run python examples/custom_fonts.py input.docx output.pdf fonts/JapaneseFont.ttf
uv run python examples/diagnostics.py \
  input.docx output.pdf diagnostics.json layout.json
```

## CLIリファレンス

すべてのオプションを使った例です。

```console
uv run docxpdf-native input.docx \
  --output output.pdf \
  --font-dir ./fonts \
  --font-substitution "MS Mincho=Noto Serif CJK JP" \
  --font-substitution "MS Gothic=Noto Sans CJK JP" \
  --strict \
  --deterministic \
  --diagnostics diagnostics.json \
  --layout-json layout.json \
  --max-pages 100
```

| オプション | 説明 |
| --- | --- |
| `-o, --output` | 出力PDF（必須）。親ディレクトリは先に作成しておいてください。 |
| `--font-dir DIRECTORY` | フォント検索ディレクトリ。複数回指定できます。 |
| `--font-substitution NAME=TARGET` | 明示的なフォント置換。複数回指定できます。 |
| `--strict` / `--lenient` | 未対応機能とフォント不足の扱い。既定は`--strict`です。 |
| `--deterministic` | 決定的なPDFメタデータと描画順を要求します。既定でも有効です。 |
| `--diagnostics PATH` | 警告、未対応機能、フォント置換、件数をJSONへ出力します。 |
| `--layout-json PATH` | PDF描画前のレイアウトモデルをJSONへ出力します。 |
| `--max-pages COUNT` | 生成できる最大ページ数。既定は10,000です。 |
| `--version` / `--help` | バージョンまたはヘルプを表示します。 |

PDF、診断JSON、レイアウトJSONには、それぞれ別のパスを指定してください。

## Python API

公開APIは4つだけです。`Converter`、`ConversionOptions`、`ConversionResult`、`FontConfiguration`をトップレベルからインポートできます。

`convert()`は入出力の形式に柔軟です。

- **入力**: `Path`、`str`、`bytes`、バイナリファイルオブジェクト
- **出力先**: `Path`、`str`、書き込み可能なバイナリファイルオブジェクト。省略しても`result.pdf_bytes`からPDFを取得できます

用途に合わせた補助メソッドもあります。

```python
converter = Converter()

# バイト列だけを扱う（Webアプリのハンドラーなどに便利です）
pdf_bytes = converter.convert_bytes(docx_bytes)

# 入出力ストリームを扱う
result = converter.convert_stream(source_stream, destination_stream)
```

`ConversionResult`には変換の記録がまとまっています。

| 属性 | 内容 |
| --- | --- |
| `page_count` | 生成されたページ数 |
| `warnings` | 変換中の警告（コード、メッセージ、位置つき） |
| `unsupported_features` | 検出された未対応機能 |
| `font_substitutions` | 行われたフォント置換（理由と位置つき） |
| `pdf_bytes` / `pdf_sha256` | PDFバイト列とそのSHA-256 |
| `layout_json` / `layout_sha256` | レイアウトモデルのJSON文字列とそのSHA-256 |
| `diagnostics` | 診断情報のPydanticモデル |

出力ファイルは同じディレクトリの一時ファイルへ書いた後に置き換えるので、変換途中の中途半端なPDFが残ることはありません。

## strictモードとlenientモード

**strictモード（既定）**は「黙って変になる」ことを許しません。検出済みの未対応機能には`UnsupportedFeatureError`、解決できないフォントには`FontNotFoundError`を送出し、PDFは作りません。明示したフォント置換はstrictモードでも使えます。

**lenientモード**は「できるところまで変換して、削ったものを正直に報告する」モードです。対応できる内容を変換し、削除した未対応機能を`ConversionResult.warnings`と診断JSONへ記録します。見つからないフォントは`default_font`（未指定ならHelvetica）へ置き換え、置換理由と段落・ラン位置を記録します。

どちらを選ぶかの目安です。

- 出力の正確さが最優先（帳票、公的文書など）→ **strict**
- 多少の欠けは許容してとにかくPDFがほしい → **lenient**（ただし診断JSONの確認をおすすめします）

なお、lenientモードでも、壊れたZIP/XML、危険なRelationship、リソース上限超過、存在しない置換先、配置不能な内容はエラーになります。安全性は緩めません。

## 困ったときは

CLIの終了コードは失敗の理由を教えてくれます。

| 終了コード | 意味 | まず試すこと |
| --- | --- | --- |
| `0` | 成功 | — |
| `2` | 入力やJSON出力の問題 | パスの存在、出力先の親ディレクトリ、JSONパスの重複を確認 |
| `3` | 未対応機能を検出 | エラーメッセージの回避策を確認するか、`--lenient`で警告として続行 |
| `4` | フォントが見つからない | `--font-dir`や`--font-substitution`で必要なフォントを用意 |
| `5` | 不正なDOCX/OOXML | ファイルが本当にDOCXか、破損していないかを確認 |
| `6` | レイアウトまたはPDF生成の失敗 | `--layout-json`で内容を確認、`--max-pages`の上限も確認 |
| `10` | 想定外の失敗 | 再現手順とともにIssueで教えてください |

よくあるつまずきどころ:

- **日本語が「□」（豆腐）になる、またはエラーコード`4`** — 日本語フォントが解決できていません。[日本語文書を変換する](#日本語文書を変換する)の設定を確認してください。
- **Wordで見た目とページ数が違う** — フォントファイルが違うことがほとんどの原因です。DOCXが要求するフォントと同じものを用意してください。
- **何が削られたのか知りたい** — `--diagnostics diagnostics.json`を付けると、未対応機能の位置と回避策がJSONで確認できます。

## 対応している機能

バージョン0.1.0では、次の範囲を実装し、単体テストまたは統合テストで確認しています。

**文書構造**

- DOCXをZIPとして読み、Content Types、Relationship、本文、スタイル、テーマ、番号定義、ヘッダー、フッター、画像、主要メタデータを解析します（`docProps/app.xml`のページ数はレイアウトに使いません）
- セクションごとのページ幅・高さ、縦横、余白、ヘッダー・フッター距離、ページ番号開始値。A4、Letterを含む任意の指定サイズをpointへ変換します
- 連続区切りと奇数・偶数ページ区切り（決定的な規則で処理します）
- ヘッダーとフッターの配置、単純な`PAGE`・`NUMPAGES`フィールドの解決

**段落とテキスト**

- 明示的な改ページ、`pageBreakBefore`、明示的な改行、タブ、空段落
- 段落前後の余白、左右・先頭行・ぶら下げインデント、固定/最小/倍数行間、`keepNext`、`keepLines`、`widowControl`
- 左・中央・右・両端揃えと自動改行（単一の可視テキストランからなる段落）
- フォント、East Asiaフォント、サイズ、太字、斜体、下線、取り消し線、色、名前付きハイライト、上付き・下付き、文字間隔、非表示テキスト
- `docDefaults`、`basedOn`、段落/文字スタイル、直接書式、テーマフォントの解決とスタイル循環の検出
- 箇条書きと番号付きリストの定義解決とラベル配置

**日本語処理**

- 横書き日本語のNFC正規化と実フォント幅による計測
- 日本語の行頭・行末禁則、英数字列の途中分割抑制
- 結合文字、Variation Selector、絵文字のZWJ列を分離しない簡易クラスタ処理
- 1ラン内でLatinとEast Asiaの指定が切り替わる文字列の、文字種ごとの描画ラン分割

**表と画像**

- 固定グリッドの表、表/列/セル幅、行高、セル余白、罫線、背景色、横/縦セル結合、セル内の上下・左右揃え
- 段落と表の順序保持、行単位の改ページ、ヘッダー行の繰り返し
- PNG/JPEGのインライン画像（EMU指定寸法と比率を保ち、本文幅へ縮小します）

**決定性**

- 内部座標は左上原点・point単位に統一し、PDF描画時だけReportLabの左下原点へ変換します
- 同じ入力、設定、フォント、依存バージョンのもとで、レイアウトJSONとPDFバイト列が再現することをテストしています

## 未対応の機能

次の機能は描画しません。strictモードでは停止し、lenientモードでは位置と回避策を含む警告を残します。

- テキストボックス、SmartArt、グラフ、埋め込みオブジェクト、OLE、マクロ、数式、WordArt
- 浮動画像と`wp:anchor`、外部リンク画像
- 縦書き、ルビ、双方向レイアウト、複雑なアラビア語シェーピング
- 変更履歴、コメント表示、コンテンツコントロール
- 複数段組、脚注、文末脚注
- `PAGE`と`NUMPAGES`以外の単純フィールド、および複雑なフィールドコード

未対応要素を検出すると、機能名、OOXMLパーツ、要素名、XPath相当の位置、対応状況、回避策を持つ`UnsupportedFeature`が作られます。

## 既知の制限

正直にお伝えします。次の点は現時点の制限です。

- Microsoft Wordのレイアウト規則を完全には再現しません。Wordで保存した見た目との一致を保証するものではありません。
- 複数のテキストランは書式を保って自動折り返ししますが、画像、タブ、明示改行が同じ段落へ複雑に混在する場合は、Wordと異なる折り返しになることがあります。
- 両端揃えは空白または日本語の分割可能位置へ余白を配ります。Word固有の均等割り付けは再現しません。
- 箇条書きと番号付きリストは番号定義を解決してラベルを描画しますが、Wordの全番号形式と複雑な階層継続は再現しません。
- 連続セクションはページ寸法と本文領域が同じで、残り領域へ収まる場合に同じページへ続けます。連続ページ内でヘッダー設定を切り替えるWord固有の挙動は再現しません。
- AutoFit表は、グリッドがなければ利用可能幅を等分する決定的な近似です。内容に基づくWordのAutoFitは再現せず、近似したことを診断警告へ記録します。
- 1行の表が本文領域より高い場合、セル内容を分割せず`LayoutError`にします。
- 複雑な文字シェーピング、フォントフォールバックを伴う1ラン内の多書体合成、縦書きは行いません。
- カスタムフォントの太字・斜体は別フェイスを自動探索しません。TrueTypeアウトラインを持たないCFF OpenTypeとTTCは、初期版のPDF埋め込み対象外です。
- フォントファイルに個々のグリフがあるかは事前検証しません。strictモードでも、指定フォント自体が見つかれば不足グリフを変換前に検出できない場合があります。
- 既知の未対応要素は検出しますが、あらゆる将来のOOXML拡張を網羅する検出器ではありません。未知のマークアップがある文書はstrictモードでも事前に停止しない場合があります。

## ページ数と決定性

ページ数は`docProps/app.xml`の値ではなく、独自レイアウトが作った`PageModel`の件数です。次の条件が変わると、ページ数や改ページ位置も変わります。

- 使用するフォントファイル、フォントの版、置換設定
- ページサイズ、余白、段落間隔、行間、表幅、画像寸法
- strict/lenient設定と、lenientで除かれる未対応内容
- `fonttools`とReportLabの版、実行環境の利用可能フォント
- AutoFit表や複数ラン折り返しなど、現在近似している機能

同じ入力、`ConversionOptions`、フォントファイル、依存バージョンを固定すると、ページ数、描画順、レイアウトJSON、PDFのSHA-256を再現することを目標にしています。異なる環境間の同一ハッシュは保証しません。

## レイアウトJSONと診断JSON

「なぜこのレイアウトになったのか」を調べたいときのために、PDFとは独立した検証用データを出力できます。

**レイアウトJSON**は、ページサイズ、ヘッダー・本文・フッター領域、ボックス座標、行、文字フラグメント、元文書の範囲を持ちます。座標は左上原点、単位はpointです。

**診断JSON**は、警告、未対応機能、フォント置換、件数のまとめです。

```console
uv run docxpdf-native input.docx -o output.pdf \
  --layout-json layout.json \
  --diagnostics diagnostics.json
```

Python APIでは、変換後の`converter.last_layout`からPydanticモデルとして取得できます。`result.layout_json`には同じレイアウトの決定的なJSON文字列、`result.diagnostics`には診断情報が入っています。

```python
from pathlib import Path

from docxpdf_native import Converter

converter = Converter()
result = converter.convert("input.docx", "output.pdf")

if result.layout_json is not None:
    Path("layout.json").write_text(result.layout_json, encoding="utf-8")
Path("diagnostics.json").write_text(
    result.diagnostics.model_dump_json(indent=2),
    encoding="utf-8",
)
```

## セキュリティ上の配慮

DOCXは信頼できない入力として扱います。公開サービスへの組み込みを想定した設計です。

- ZIPをディスクへ展開せず、危険なパス、重複パーツ、暗号化パーツ、ZIP内シンボリックリンクを拒否します。
- 圧縮前DOCXサイズ、合計展開サイズ、XMLパーツ、画像、XML深度、段落、ラン、表セル、ページに上限を設けます。
- DTDとエンティティ宣言を拒否し、XML外部実体を使いません。
- Relationshipの重複、欠落先、パッケージ外参照、循環を検査します。外部Relationshipは既定で無効です。外部画像は未対応として記録し、取得しません。
- 入力と出力が同じファイルまたはハードリンクの場合、存在しない出力親ディレクトリ、出力先と親ディレクトリのシンボリックリンクを拒否します。

既定のリソース上限は、DOCX 100 MiB、合計展開512 MiB、XMLパーツ32 MiB、画像64 MiB、XML深度128、段落100,000、ラン1,000,000、表セル100,000、ページ10,000です。`ConversionOptions.resource_limits`でさらに小さくできます。

ただし、リソース上限はサービス全体のCPU時間やメモリ使用量を保証するものではありません。公開サービスへ組み込む場合は、プロセス単位の時間・メモリ制限も併用してください。

## 開発に参加する

開発環境の準備から検証まで、必要なコマンドはこれだけです。

```console
uv sync
uv run pytest
uv run pytest --cov=docxpdf_native --cov-report=term-missing
uv run ruff check .
uv run ruff format --check .
uv run mypy src
```

まとめて確認したいときは`make check`が便利です（テスト、カバレッジ、lint、フォーマット、型チェックを順に実行します）。テストは外部ネットワーク、Microsoft Office、LibreOfficeを必要としません。

## 公式PDF回帰ベンチマーク

`benchmark/data/`には、デジタル庁が同じ項目で公開しているDOCX/PDFペア3組と、版一致の根拠、サイズ、SHA-256、取得日を記録した`reference-manifest.json`があります。公式PDFを正解として、候補PDFを3回生成し、ページ数、ページ枠、回転、正規化全文、ページ別テキスト、改ページ境界、PDF/レイアウトJSONの再現性を比較できます。

```console
uv run python -m benchmark.validate_reference
# または
make benchmark
```

macOSで`/System/Library/Fonts/Supplemental/Arial Unicode.ttf`が存在する場合、ベンチマークはそれを固定代替フォントとして使います（フォントはリポジトリへコピーしません）。それ以外の環境では、`BenchmarkConfig(font_path=...)`へ再配布・埋め込み条件を確認した日本語TTF/OTFを明示してください。フォントが違えば結果は比較できません。

最新の実測レポートは`benchmark/results/report.md`、機械可読結果は`results.json`、全ページ画像の目視結果は`visual-validation.md`にあります。現在の3文書では候補PDFの3回再現性は確認できましたが、公式PDFとのページ数・改ページ・視覚レイアウトはまだ合致していません。ここは今後の改善ポイントとして正直に記録しています。

ベンチマーク本体は外部プロセスを起動しません。画像検証は変換後の正解PDFと候補PDFを、同じPoppler設定でリポジトリ外から直接画像化して行います。

## ライセンス

MIT Licenseです。全文は[`LICENSE`](LICENSE)を参照してください。

変更履歴は[`CHANGELOG.md`](CHANGELOG.md)にあります。質問や不具合報告、機能の要望はIssueでお気軽にどうぞ。
