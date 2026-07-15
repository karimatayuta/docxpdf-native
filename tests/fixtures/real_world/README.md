# 実世界DOCXコーパス

`tests/integration/test_real_world_corpus.py` が使う、ウェブ上で公開されている実世界のDOCXファイル群です。古い世代のWord(2007〜2013、互換モード、Kingsoft製など)で作成され、埋め込みOLEオブジェクト・フィールドコード・VML画像・変更履歴・脚注などの「難しい」要素を含む文書を集めています。

第三者の文書を再配布しないため、**DOCXファイル本体はリポジトリにコミットされていません**(`.gitignore` 対象)。メタデータ(取得元URL、SHA-256、参照ページ数、特徴)は `manifest.json` に記録されています。

## 再取得方法

`manifest.json` の各エントリの `source_url` からダウンロードし、`filename` の名前でこのディレクトリに保存してください。

```console
python3 - <<'EOF'
import hashlib
import json
import pathlib
import urllib.request

directory = pathlib.Path("tests/fixtures/real_world")
manifest = json.loads((directory / "manifest.json").read_text(encoding="utf-8"))
for entry in manifest["files"]:
    target = directory / entry["filename"]
    if target.exists():
        continue
    print("fetching", entry["filename"])
    request = urllib.request.Request(entry["source_url"], headers={"User-Agent": "Mozilla/5.0"})
    target.write_bytes(urllib.request.urlopen(request).read())
    digest = hashlib.sha256(target.read_bytes()).hexdigest()
    if digest != entry["sha256"]:
        print("  WARNING: SHA-256 mismatch (upstream file may have changed)")
EOF
```

ファイルが存在しない場合、コーパステストは自動的にスキップされます。取得元が消えた場合はSHA-256の一致する代替ミラー(web.archive.org など)を探してください。

## ページ数の参照値について

`app_pages` は各DOCXの `docProps/app.xml` に記録されたページ数(最後に保存したWord環境でのページネーション結果)です。`app_pages_reliable: false` が付いたエントリは、この値が信頼できないと判明したもので(理由は `note` を参照)、ページ数の検証からは除外され、エラーなし変換のみ検証されます。
