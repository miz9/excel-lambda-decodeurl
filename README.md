[README.md](https://github.com/user-attachments/files/34163115/README.md)
# Excel LAMBDA `_DECODEURL`

Excel の数式だけで、パーセントエンコードされた URL 文字列をデコードするユーザー定義 LAMBDA 関数です。VBA、Office Scripts、外部アドインを使用しません。

An Excel LAMBDA function that decodes percent-encoded URL strings using formulas only. No VBA, Office Scripts, or external add-ins are required.

## Features

- Excel の LAMBDA 関数として利用可能
- `%xx` 形式のパーセントエンコードをデコード
- UTF-8 の 2～4 バイト文字を処理
- VBA・マクロ不要
- ワークシート関数として再利用可能

## Requirements

次の関数を利用できる Microsoft Excel が必要です。

- `LAMBDA`
- `LET`
- `XLOOKUP`
- `BITAND`
- `BITLSHIFT`
- `HEX2DEC`

Microsoft 365 または、これらの関数をサポートする Excel 環境を使用してください。

## Installation

1. リポジトリ内の [`_DECODEURL V1.44`](./_DECODEURL%20V1.44) を開きます。
2. 数式全体をコピーします。
3. Excel で **数式** → **名前の管理** → **新規作成** を開きます。
4. 名前に `_DECODEURL` を入力します。
5. **参照範囲**へコピーした LAMBDA 数式を貼り付けます。
6. 保存して、ワークシートから `_DECODEURL` を呼び出します。

## Usage

```excel
=_DECODEURL(A1)
```

直接文字列を指定する例:

```excel
=_DECODEURL("Excel%20LAMBDA")
```

期待される結果:

```text
Excel LAMBDA
```

UTF-8でエンコードされた日本語の例:

```excel
=_DECODEURL("%E6%97%A5%E6%9C%AC%E8%AA%9E")
```

期待される結果:

```text
日本語
```

## Version

Current version: **V1.44**

## Notes

- 入力値は文字列として扱われます。
- URL全体だけでなく、パーセントエンコードされた文字列にも使用できます。
- Excel のバージョンや更新チャネルによって、必要な関数を利用できない場合があります。

## Repository topics

`excel` `excel-lambda` `excel-formulas` `microsoft-excel` `lambda` `url-decoding` `decode-urls`

## License

ライセンスは未指定です。再利用条件を明確にする場合は、別途 `LICENSE` ファイルを追加してください。

---

Created by [miz9](https://github.com/miz9)
