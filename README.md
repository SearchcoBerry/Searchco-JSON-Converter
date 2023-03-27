# search-co-json
さちこの時間割関連のスクリプト

## 時間割をjsonに変換

### 設定ファイルを作成
.envファイルを作成する。以下のコマンドでenvファイルのサンプルを出力できる。

```
$ cp .env.example .env
```

#### 設定項目の説明

```
IMPORT_PATH=エクセルのファイルのパス
EXPORT_PATH=出力するjsonファイルのパス
MIN_COL=探索するセルの最小列を指定(基本弄らない)
MAX_COL=探索するセルの最大列を指定(基本弄らない)
MIN_ROW=探索するセルの最大行を指定(基本弄らない)
```

### コマンドを入力

このREADMEと同じ階層で以下のコマンドを入力。

```
$ make timetable-to-json
```