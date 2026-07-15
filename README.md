# docxpdf-native

**DOCXを、Pythonだけで、そのままPDFに。**

`docxpdf-native`は、DOCXの中身（Office Open XML）を直接読み取り、独自にページ配置を計算してPDFを生成するPythonライブラリです。Microsoft WordもLibreOfficeもブラウザーも外部レンダラーもクラウドAPIも呼び出しません。変換中のネットワーク接続も一切不要です。

「Wordをサーバーに入れられない」「変換のたびにネットへ出したくない」「同じ入力からいつも同じPDFがほしい」——そんな場面のために作られています。

ひとつだけ、最初にお伝えしたい大切なことがあります。このライブラリは**仕様を限定したDOCXのサブセット**を対象としており、Microsoft Wordとの完全互換は保証しません。既定の**lenientモード**では、対応できる内容は本来の見た目のまま描画し、描画しきれない部分（埋め込みオブジェクトのプレビュー、EMF/WMFなど変換できない画像形式など）は元の寸法を保ったプレースホルダーへ置き換えて、変換を止めずに警告として記録します。出力の正確さを優先したい場合は、明示的に**strictモード**を選べば、内容そのものが失われる機能を検出した時点で正直に停止します。「何ができて、何ができないか」を隠さないことが、このライブラリの設計方針です。

## 目次

- [特徴](#特徴)
- [動作要件](#動作要件)
- [インストール](#インストール)
- [はじめての変換](#はじめての変換)
- [日本語文書を変換する](#日本語文書を変換する)
- [同梱フォントとCalibri/Cambriaの自動置換](#同梱フォントとcalibricambriaの自動置換)
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

GIF/BMP/TIFFなどPNG/JPEG以外の画像形式も描画したい場合だけ、任意で[`Pillow`](https://python-pillow.org/)を追加してください（`uv add "docxpdf-native[images]"`）。無くても変換は止まらず、対象の画像は同サイズのプレースホルダーになります。

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

**フォント設定なしで、そのまま変換できます。** 日本語のDOCXが「MS 明朝」「MS ゴシック」「游明朝」「游ゴシック」「メイリオ」といった、変換環境には存在しないことがほとんどの定番フォントを要求してきても、docxpdf-nativeはOSにインストール済みの日本語フォント（macOSならヒラギノ明朝 ProN/ヒラギノ角ゴシックなど、Windowsなら游明朝/游ゴシック/メイリオなど、Linuxなら[Noto CJK](https://github.com/notofonts/noto-cjk)/IPAフォントなど）を自動検出し、明朝系はserif、ゴシック系はsansへと決定的に置き換えます。全角・半角の表記ゆれ（「ＭＳ　ゴシック」のような全角MS・全角スペース）も正規化して同じ置換先に解決します。置換は`result.font_substitutions`（`reason="builtin-cjk"`）にそのまま記録されるので、何が起きたかは常に確認できます。

```console
uv run docxpdf-native input.docx --output output.pdf
```

```python
from docxpdf_native import Converter

result = Converter().convert("input.docx", "output.pdf")
for substitution in result.font_substitutions:
    print(substitution.requested_font, "->", substitution.selected_font, substitution.reason)
```

この組み込み置換はstrictモードでも適用されます（フォント不足で変換が止まることはありません）。CJKフォントが1つも検出できない環境でだけ、lenientモードの最終フォールバックとしてHelveticaが使われ、豆腐（□）になる可能性があることが警告ログに残ります。詳しくは[困ったときは](#困ったときは)を参照してください。

### フォントファイルを自分で用意する

自動検出はOSごとに異なるフォントを使うため、**同じ入力でも実行環境が変わるとページ数やPDFバイト列が変わる**ことがあります。CI、サーバー、複数人での開発など、環境をまたいで結果を完全に一致させたい場合は、従来どおりフォントファイルを自分で用意し、明示的に指定してください。

1. 配布・埋め込み条件を確認した日本語TTF/OTFを入手します。例えば[Noto CJK](https://github.com/notofonts/noto-cjk/releases)のリリースページから`NotoSansCJKjp-Regular.otf`（ゴシック体）や`NotoSerifCJKjp-Regular.otf`（明朝体）をダウンロードできます。
2. リポジトリのルート（またはプロジェクトの作業ディレクトリ）に`fonts/`ディレクトリを作成し、ダウンロードしたフォントファイルを置きます。

   ```console
   mkdir fonts
   # ダウンロードしたフォントファイルを fonts/ に配置
   ```

   **`fonts/`ディレクトリが存在しない状態で`font_directories`や`--font-dir`に指定するとエラーになります**（CLIは終了コード`2`、Python APIは`ValueError`）。必ず先にディレクトリを作成し、ファイルを配置してから指定してください。
3. `FontConfiguration`または`--font-dir`/`--font-substitution`で明示的に指定します。

```python
from pathlib import Path

from docxpdf_native import ConversionOptions, Converter, FontConfiguration

fonts = FontConfiguration(
    registered_fonts={
        "Noto Sans CJK JP": Path("fonts/NotoSansCJKjp-Regular.otf"),
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

- **フォントの検索順**は、APIで登録したフォント（`registered_fonts`）、`font_directories`、環境変数、OS標準フォントディレクトリ、[同梱フォント](#同梱フォントとcalibricambriaの自動置換)、ReportLab標準フォント、明示的な置換（`substitutions`）、組み込みの日本語フォント置換マップ、組み込みのメトリック互換置換マップ（Calibri/Cambria）、既定フォント（lenientモードのみ）の順です。2つの組み込み置換マップはstrictモードでも適用されます。
- 同じファミリー名を共有する複数のフェイス（Regular/Bold/Italic等）がスキャンで見つかった場合、素のファミリー名はRegularフェイスに解決されます。スタイル付きフェイスは「Times New Roman Bold」のようなフルネームで引き続き解決できます。
- 環境変数の既定名は`DOCXPDF_NATIVE_FONT_DIRS`で、複数のパスはOSのパス区切り文字（macOS/Linuxは`:`）で区切ります。
- 特定のフォントファイルに別名を付ける`registered_fonts`はPython API専用です。CLIの`--font-dir`はディレクトリ検索の追加だけを行います。`registered_fonts`に`.ttc`/`.otc`（TrueType/OpenType Collection）を指定した場合は、埋め込み可能な最初のフェイスが自動的に選ばれます。
- 環境間で結果を完全に一致させたい場合は、`include_system_fonts=False`にしてOSのフォントに依存しない構成にするのがおすすめです（`examples/custom_fonts.py`がこの構成です）。組み込みの日本語フォント置換マップも、システムフォントが見つからない限り実際には置換できないため、`include_system_fonts=False`にした環境では`substitutions`や`default_font`を明示してください。
- `.ttc`/`.otc`はフェイス単位でスキャンし、フェイスごとに正しいものをPDFへ埋め込みます。ただしTrueTypeアウトライン（`glyf`テーブル）を持たないCFFベースのフェイス（macOSのヒラギノ収録フォントなど）はReportLabで埋め込めないため、索引から除外されます。除外されたフェイスしか含まないファイルは、フォントが見つからなかったものとして扱われます。
- lenientモードの最終フォールバックは、要求フォントがCJK系（`east_asia`指定、組み込みマップに一致、またはフォント名自体にCJK文字を含む）であれば検出済みのシステムCJKフォントを使い、そうでなければHelveticaを使います。CJKフォントが1つも検出できない環境でだけ、CJK系の要求もHelveticaへ落ち、日本語グリフを持たないため豆腐（□）になります（警告ログが出力されます）。

`examples/`にそのまま実行できる例を用意しています。

```console
uv run python examples/basic_conversion.py input.docx output.pdf
uv run python examples/custom_fonts.py input.docx output.pdf fonts/JapaneseFont.ttf
uv run python examples/diagnostics.py \
  input.docx output.pdf diagnostics.json layout.json
```

## 同梱フォントとCalibri/Cambriaの自動置換

Word 2007以降の既定フォントである**Calibri**（と見出し用の**Cambria**）はプロプライエタリなフォントで、変換環境にはまず入っていません。字幅の異なるフォント（Helvetica等）へ置き換えると折り返し行数が変わり、ページ数が大きくずれます。

このためdocxpdf-nativeは、**メトリック互換**（同じ字幅）のオープンソースフォントを2ファミリー同梱しています。

| 要求フォント | 置換先（同梱） | ライセンス | 入手元 |
| --- | --- | --- | --- |
| Calibri（Light/Bold/Italic等の派生名を含む） | Carlito | SIL OFL 1.1 | [googlefonts/carlito](https://github.com/googlefonts/carlito) |
| Cambria（Math/Bold/Italic等の派生名を含む） | Caladea | SIL OFL 1.1 | [huertatipografica/Caladea](https://github.com/huertatipografica/Caladea) |

- 置換は設定なしで自動的に行われ、`result.font_substitutions`に`reason="builtin-metric-compatible"`として記録されます。strictモードでも適用されます（変換は止まりません）。
- CalibriやCambriaそのものがシステムにインストールされている場合は、実フォントがそのまま使われ、置換は行われません。ユーザー明示の`substitutions`も組み込み置換より優先されます。
- 同梱フォントは検索順で「OS標準フォントディレクトリの後、ReportLab標準フォントの前」に位置します。「Carlito」「Caladea」を直接要求することもできます（Regular/Bold/Italic/Bold Italicの4フェイスを収録）。
- フォントファイルは`src/docxpdf_native/fonts/data/`にOFL 1.1のライセンス文書（`OFL.txt`）と取得元・SHA-256の記録（`SOURCES.txt`）とともに同梱されており、本体のMITライセンスとは別のライセンスで配布されます。ネットワーク接続は一切不要です。

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
| `--strict` / `--lenient` | 未対応機能とフォント不足の扱い。既定は`--lenient`です。 |
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

**lenientモード（既定）**は「古いDOCXも含め、とにかく最後まで変換する」モードです。コンテンツコントロール（`w:sdt`）、変更履歴（`w:ins`/`w:del`/`w:moveFrom`/`w:moveTo`）、複合フィールド（`w:fldChar`+`w:instrText`）、VML画像（`w:pict`、`w:object`内の`v:shape`+`v:imagedata`）といった、実世界の古いDOCXに頻出する構造は本来の内容のまま描画します。それでも描画できない要素（埋め込みオブジェクトのプレビュー、EMF/WMFなど変換できない画像形式、外部リンク画像、テキストボックス/SmartArt/グラフなど寸法しか分からない図形）は、元の寸法を保った同サイズのプレースホルダー枠に置き換え、ページ数とレイアウト全体のずれを防ぎます。何が置き換えられたかは、コード`content_placeholder`の`ConversionResult.warnings`と診断JSONに、位置・要素名・回避策つきで記録されます。見つからないフォントは`default_font`（未指定ならHelvetica）へ置き換え、置換理由と段落・ラン位置を記録します。

**strictモード**は「黙って変になる」ことを一切許しません。テキストボックス内のテキスト、SmartArt、グラフ、数式、脚注・文末脚注、コメント、縦書き、ルビ、双方向レイアウト、複雑なアラビア語シェーピング、複数段組、マクロなど、**内容そのものが失われる**機能を検出した時点で`UnsupportedFeatureError`を、解決できないフォントには`FontNotFoundError`を送出し、PDFは作りません。同サイズのプレースホルダーで代替できる構造（VML画像、埋め込みオブジェクトのプレビュー、コンテンツコントロール、変更履歴、複合フィールドなど）は、strictモードでもlenientモードと同じように描画・代替されます。これらはもう「未対応」として停止する対象ではなく、常に報告されたうえで変換が続く一つの挙動だからです。

どちらを選ぶかの目安です。

- 出力の正確さが最優先（帳票、公的文書など）→ **strict**
- 古いWordファイルも含め、とにかく変換したい → **lenient**（既定。ただし診断JSONの確認をおすすめします）

なお、lenientモードでも、壊れたZIP/XML、危険なRelationship、リソース上限超過、存在しない置換先、配置不能な内容はエラーになります。安全性は緩めません。

## 困ったときは

CLIの終了コードは失敗の理由を教えてくれます。

| 終了コード | 意味 | まず試すこと |
| --- | --- | --- |
| `0` | 成功 | — |
| `2` | 入力やJSON出力の問題 | パスの存在、出力先の親ディレクトリ、JSONパスの重複を確認 |
| `3` | 未対応機能を検出 | エラーメッセージの回避策を確認するか、`--lenient`で警告として続行 |
| `4` | フォントが見つからない | 通常は日本語フォントの組み込み置換マップが解決するため稀です。カスタムフォント名が原因のことが多いので、`--font-dir`や`--font-substitution`で必要なフォントを用意してください |
| `5` | 不正なDOCX/OOXML | ファイルが本当にDOCXか、破損していないかを確認 |
| `6` | レイアウトまたはPDF生成の失敗 | `--layout-json`で内容を確認、`--max-pages`の上限も確認 |
| `10` | 想定外の失敗 | 再現手順とともにIssueで教えてください |

よくあるつまずきどころ:

- **日本語が「□」（豆腐）になる** — システムに日本語フォントが1つも見つからない環境でだけ起こります。`result.font_substitutions`（または`--diagnostics`）を確認し、`reason`が`builtin-cjk`や`lenient-default`になっている置換が実在の日本語フォントを指しているか確認してください。見つからない場合は[フォントファイルを自分で用意する](#フォントファイルを自分で用意する)の手順でフォントを配置してください。
- **エラーコード`4`（フォントが見つからない）** — 日本語の定番フォント名（MS明朝/MSゴシック/游明朝/游ゴシック/メイリオなど）は組み込み置換マップが自動的に解決するため、通常はここで止まりません。それ以外のカスタムフォント名や、CFFベースの`.ttc`しか無くPDFに埋め込めない場合に発生します。`--font-dir`や`--font-substitution`で必要なフォントを用意してください。
- **Wordで見た目とページ数が違う** — フォントファイルが違うことがほとんどの原因です。自動検出されるシステムフォントは実行環境ごとに異なるため、環境間で完全に一致させたい場合は[フォントファイルを自分で用意する](#フォントファイルを自分で用意する)を参照してください。
- **何が削られたのか知りたい** — `--diagnostics diagnostics.json`を付けると、未対応機能とフォント置換の位置・理由がJSONで確認できます。

## 対応している機能

バージョン0.1.0では、次の範囲を実装し、単体テストまたは統合テストで確認しています。

**文書構造**

- DOCXをZIPとして読み、Content Types、Relationship、本文、スタイル、テーマ、番号定義、ヘッダー、フッター、画像、主要メタデータを解析します（`docProps/app.xml`のページ数はレイアウトに使いません）
- セクションごとのページ幅・高さ、縦横、余白、ヘッダー・フッター距離、ページ番号開始値。A4、Letterを含む任意の指定サイズをpointへ変換します
- 連続区切りと奇数・偶数ページ区切り（決定的な規則で処理します）
- ヘッダーとフッターの配置、単純な`PAGE`・`NUMPAGES`フィールドの解決
- コンテンツコントロール（`w:sdt`）は、ブロックレベル（段落・表を包む）・インラインレベル（ラン・ハイパーリンクを包む）のどちらも、入れ子になっていても中身をそのまま展開して描画します
- `mc:AlternateContent`は`mc:Fallback`（多くは互換性の高いVML表現）を優先して処理し、`mc:Fallback`が無い場合だけ`mc:Choice`の内容を使います

**段落とテキスト**

- 明示的な改ページ、`pageBreakBefore`、明示的な改行、タブ、空段落
- 段落前後の余白、左右・先頭行・ぶら下げインデント、固定/最小/倍数行間、`keepNext`、`keepLines`、`widowControl`
- 隣接する段落間の余白は、Wordと同様に前段落の後余白と次段落の前余白を加算せず大きい方だけを適用します（表やページ区切りをまたぐ場合はリセットされます）
- セクションの行グリッド（`w:docGrid`の`type="lines"`・`"linesAndChars"`）に対応し、本文段落と表セル内の行高さをグリッド（`w:linePitch`）の倍数へ切り上げます。日本語の既定テンプレートで使われる設定で、対応しないと折り返し行が詰まりすぎてページ数を過小評価します。固定行間（`exact`）の段落と、縦方向に結合されたセルは対象外です
- 左・中央・右・両端揃えと自動改行（単一の可視テキストランからなる段落）
- フォント、East Asiaフォント、サイズ、太字、斜体、下線、取り消し線、色、名前付きハイライト、上付き・下付き、文字間隔、非表示テキスト
- `docDefaults`、`basedOn`、段落/文字スタイル、直接書式、テーマフォントの解決とスタイル循環の検出
- 箇条書きと番号付きリストの定義解決とラベル配置
- 変更履歴は承認済みの見た目で描画します（挿入`w:ins`・移動先`w:moveTo`は表示、削除`w:del`・移動元`w:moveFrom`は非表示）
- 複合フィールド（`w:fldChar`のbegin/separate/end + `w:instrText`）は、DOCX保存時にキャッシュされた結果テキストをそのまま描画します。ネストしたフィールドは最外層の結果だけを描画し、内側の命令・結果は破棄します。命令が`PAGE`・`NUMPAGES`の場合は`w:fldSimple`と同じくページ番号センチネルに解決します

**日本語処理**

- 横書き日本語のNFC正規化と実フォント幅による計測
- 日本語の行頭・行末禁則、英数字列の途中分割抑制
- 結合文字、Variation Selector、絵文字のZWJ列を分離しない簡易クラスタ処理
- 1ラン内でLatinとEast Asiaの指定が切り替わる文字列の、文字種ごとの描画ラン分割

**表と画像**

- 固定グリッドの表、表/列/セル幅、行高、セル余白、罫線、背景色、横/縦セル結合、セル内の上下・左右揃え
- 段落と表の順序保持、行単位の改ページ、ヘッダー行の繰り返し
- 表がページ途中（先行する段落の続きなど）から始まる場合も、現在ページの残り高さに収まる行から詰めて配置します（縦方向に結合されたセルを含む表は、結合セルが分割境界をまたがないよう、安全側で表全体を新しいページへ送ります）
- PNG/JPEGのインライン画像（`wp:inline`・`wp:anchor`ともに対応。EMU指定寸法と比率を保ち、本文幅へ縮小します）
- Pillow（`images`エクストラ、`uv add "docxpdf-native[images]"`）がインストールされていれば、GIF/BMP/TIFFなど他形式の画像もPNGへ自動変換して描画します。未インストールの場合は同サイズのプレースホルダーになります
- VML画像（`w:pict`、および`w:object`（埋め込みオブジェクト）内の`v:shape`+`v:imagedata`）。寸法はVMLの`style`属性（pt/in/cm/mm/px）、または`w:object`の`w:dxaOrig`/`w:dyaOrig`（存在する場合）から取得します
- 描画できない画像・図形・オブジェクト（EMF/WMFプレビュー、変換できない画像形式、外部リンク画像、テキストボックス/SmartArt/グラフ/WordArtなど寸法だけ判明する要素）は、元の寸法を保った同サイズのプレースホルダー枠（薄いグレーの矩形とラベル）として配置し、ページ数のずれを防ぎます

**決定性**

- 内部座標は左上原点・point単位に統一し、PDF描画時だけReportLabの左下原点へ変換します
- 同じ入力、設定、フォント、依存バージョンのもとで、レイアウトJSONとPDFバイト列が再現することをテストしています

## 未対応の機能

次の内容は、その中身をそのままの形では描画しません。strictモードでは検出した時点で停止し、lenientモード（既定）では位置と回避策を含む警告を残したうえで変換を続けます。

- テキストボックス、SmartArt、グラフの**中身**（文字・データ・意匠）、数式、WordArtの**文字そのもの**（枠自体は同寸プレースホルダーとして保持します）
- コメントの内容（コメントの付与位置情報は無視し、本文はそのまま変換します）
- 脚注・文末脚注の本文
- 縦書き、ルビ、双方向レイアウト（`bidi`/`rtl`）、複雑なアラビア語シェーピング
- 複数段組（`w:cols`が2段以上）
- マクロ（`vbaProject.bin`の存在を検出して記録します。実行・展開はしません）

未対応要素を検出すると、機能名、OOXMLパーツ、要素名、XPath相当の位置、対応状況、回避策を持つ`UnsupportedFeature`が作られます。

これとは別に、**内容ではなく見た目だけ**を代替した箇所（埋め込みオブジェクトのプレビュー、EMF/WMFなど変換できない画像形式、外部リンク画像、寸法しか分からない浮動図形など）も、同じ`UnsupportedFeature`の形（コード`content_placeholder`、`status="placeholder"`）で診断に記録されます。ただしこちらはstrict/lenientどちらのモードでも変換を止めません。ページ数とレイアウトを保つことを優先し、常に警告として報告する設計です。

## 既知の制限

正直にお伝えします。次の点は現時点の制限です。

- Microsoft Wordのレイアウト規則を完全には再現しません。Wordで保存した見た目との一致を保証するものではありません。
- 複数のテキストランは書式を保って自動折り返ししますが、画像、タブ、明示改行が同じ段落へ複雑に混在する場合は、Wordと異なる折り返しになることがあります。
- 両端揃えは空白または日本語の分割可能位置へ余白を配ります。Word固有の均等割り付けは再現しません。
- 箇条書きと番号付きリストは番号定義を解決してラベルを描画しますが、Wordの全番号形式と複雑な階層継続は再現しません。
- 連続セクションはページ寸法と本文領域が同じで、残り領域へ収まる場合に同じページへ続けます。連続ページ内でヘッダー設定を切り替えるWord固有の挙動は再現しません。
- AutoFit表は、グリッドがなければ利用可能幅を等分する決定的な近似です。内容に基づくWordのAutoFitは再現せず、近似したことを診断警告へ記録します。
- 表の行が現在ページの残り高さに収まらない場合、Wordの既定（行の途中での改ページを許可）にならい、その行を現在ページの残り高さで分割して複数ページにまたがって描画します。分割位置は決定的な近似で、Word同様に分割位置へ罫線が入りますが、Wordの実際の改行位置と厳密には一致しません。画像やプレースホルダーはブロック単位のため分割できず、どちらか一方のページへまるごと配置します。現在ページの残りに行の最初の1行分すら収まらない場合は、分割せず次ページへ丸ごと送ります。`w:cantSplit`指定の行と、縦方向に結合されたセルを含む行は分割せず、まず次ページへ丸ごと送ることを優先します（前者は本文領域以下である限り分割しません。行単体で本文領域より高い場合は`cantSplit`でも分割にフォールバックします〔`table_row_split`診断警告〕。後者はページ境界をまたぐ場合、従来どおり`LayoutError`にします）。
- 複雑な文字シェーピング、フォントフォールバックを伴う1ラン内の多書体合成、縦書きは行いません。
- カスタムフォントの太字・斜体は別フェイスを自動探索しません。`.ttc`/`.otc`はTrueTypeアウトライン（`glyf`テーブル）を持つフェイスのみ対応しており、CFFベースのフェイス（macOSのヒラギノ収録フォントなど）はPDF埋め込み対象外です。
- フォントファイルに個々のグリフがあるかは事前検証しません。strictモードでも、指定フォント自体が見つかれば不足グリフを変換前に検出できない場合があります。
- 既知の未対応要素は検出しますが、あらゆる将来のOOXML拡張を網羅する検出器ではありません。未知のマークアップがある文書はstrictモードでも事前に停止しない場合があります。
- 表のセルに入れ子になった表（セル内`w:tbl`）は展開しません。セル内の段落だけを読み取ります。
- 複合フィールド（`w:fldChar`）は段落単位で追跡します。目次（TOC）フィールドのように結果が複数段落にまたがる場合、段落をまたいだ時点でフィールドの追跡を打ち切り、それ以降は通常の段落として描画します（内容は失われませんが、フィールドとしての特別な扱いはそこで終わります）。
- VML図形グループ（`v:group`）の中に複数の`v:shape`がある場合、最初に見つかった1つだけを画像またはプレースホルダーとして扱います。
- `mc:AlternateContent`は常に`mc:Fallback`を優先します。`mc:Choice`側にしか無い情報（新しい描画効果など）は、対応するFallback表現に含まれていない限り再現しません。

## ページ数と決定性

ページ数は`docProps/app.xml`の値ではなく、独自レイアウトが作った`PageModel`の件数です。次の条件が変わると、ページ数や改ページ位置も変わります。

- 使用するフォントファイル、フォントの版、置換設定
- ページサイズ、余白、段落間隔、行間、表幅、画像寸法
- strict/lenient設定と、strictモードで停止対象になる内容の有無
- 実行環境にPillow（`images`エクストラ）があるかどうか（無い場合はGIF/BMP/TIFF等がプレースホルダーへ置き換わります）
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
- Relationshipの重複、欠落先、パッケージ外参照、循環を検査します。外部リンク（外部ハイパーリンク、外部画像など）は、実世界のDOCXにほぼ必ず含まれるため存在自体では拒否しません。ただし外部ターゲットのバイト列は一切取得しません。ハイパーリンクはリンク先に関係なくテキストだけ描画し、外部画像は宣言された寸法が分かれば同サイズのプレースホルダーとして記録し、変換は続行します。
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
