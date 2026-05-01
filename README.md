# 交通費・定期申請書作成システム

## 概要

本プロジェクトは、Webフォームから入力された申請情報をもとに、Excel形式の申請書を自動生成するシステムです。

主に以下の2種類の帳票作成に対応しています。

- 定期申請書
- 交通費精算書

利用者はHTMLフォームから必要事項を入力し、必要に応じて定期券の写真や前月分の定期申請Excelを添付します。  
送信された内容をAWS Lambda上のJavaプログラムで処理し、S3上のExcelテンプレートへ値を転記したうえで、完成したExcelファイルを生成します。

## 主な機能

### 定期申請書作成

定期申請フォームから以下の情報を入力し、定期申請書Excelを生成します。

- 提出日
- 勤務地名
- 勤務先住所
- 申請者氏名
- 届出の理由
- 上記理由が生じた年月日
- 交通機関名
- 移動区間
- 所要時間
- 購入期間
- 定期代金額
- 備考
- 利用開始日
- 定期券の写真

また、前月分の定期申請Excelを添付した場合は、前月Excelから勤務先情報や定期経路情報を読み取り、今回申請に流用できます。

### 交通費精算書作成

交通費精算フォームから以下の情報を入力し、交通費精算書Excelを生成します。

- 申請年月
- 社員No
- 申請者氏名
- プロジェクト名
- 月日
- 行き先
- 利用路線
- 理由
- 出発
- 到着
- 単価
- 片道／往復
- 交通費区分

単価と片道／往復の選択に応じて、画面上で金額と合計金額を自動計算します。

### ファイルアップロード

画像ファイルや前月分Excelファイルは、ブラウザから直接S3へアップロードします。

アップロード時は、まずLambdaから署名付きPUT URLを取得し、そのURLに対してブラウザがファイル本体をPUTします。

このため、ファイル本体はAPI GatewayやLambdaのリクエスト本文には直接含まれません。  
Lambdaには、S3上の `file_key` のみが送信されます。

### 一時ファイル削除

アップロードされた一時ファイルは、Excel生成処理後にJava側で削除します。

対象となる一時ファイルは主に以下です。

- 定期券写真
- 前月分の定期申請Excel

削除対象は `TEMP_PREFIX` 配下のS3オブジェクトに限定し、テンプレートファイルや生成済みファイルを誤って削除しないようにしています。

## システム構成

```text
ブラウザ
  |
  | 1. 入力内容送信
  | 2. ファイルアップロード用URL取得
  v
API Gateway
  |
  v
AWS Lambda Java
  |
  | S3テンプレート取得
  | 添付画像取得
  | 前月Excel取得
  | Excel生成
  v
Amazon S3
```

## 処理の流れ

### 通常の定期申請書作成

```text
1. ユーザーがHTMLフォームに申請情報を入力
2. 定期券写真がある場合、S3へアップロード
3. HTMLから /generate APIへJSONを送信
4. LambdaがS3からExcelテンプレートを取得
5. TeikiExcelServiceがテンプレートへ入力内容を転記
6. 完成したExcelをS3へアップロード
7. Lambdaが署名付きダウンロードURLを返却
8. ブラウザがExcelファイルをダウンロード
```

### 前月Excelを利用する定期申請書作成

```text
1. ユーザーが前月分の定期申請Excelを選択
2. 前月ExcelをS3へアップロード
3. HTML側で一部入力項目を非表示化
4. HTMLから /generate APIへJSONを送信
5. Lambdaが前月ExcelをS3から /tmp へダウンロード
6. TeikiExcelReadServiceが前月Excelの内容を読み取り
7. 今回入力された提出日・理由発生日・利用開始日・写真情報と結合
8. TeikiExcelServiceが新しい定期申請書を生成
9. 一時アップロードファイルをS3から削除
10. 完成したExcelのダウンロードURLを返却
```

### 交通費精算書作成

```text
1. ユーザーが交通費精算フォームに入力
2. HTML側で明細件数と合計金額を計算
3. HTMLから /carfare_generate APIへJSONを送信
4. Lambdaが交通費精算書テンプレートをS3から取得
5. KotsuhiExcelServiceがテンプレートへ入力内容を転記
6. 完成したExcelをS3へアップロード
7. Lambdaが署名付きダウンロードURLを返却
8. ブラウザがExcelファイルをダウンロード
```

## ディレクトリ・ファイル構成

```text
project-root/
├── request.html
├── Env.java
├── LambdaHandler.java
├── S3Service.java
├── TeikiExcelService.java
├── TeikiExcelReadService.java
├── KotsuhiExcelService.java
└── ApiResponse.java
```

## ファイル説明

### request.html

定期申請書作成用のHTMLフォームです。

主な役割は以下です。

- 申請情報の入力画面表示
- 定期明細行の追加・削除
- 所要時間と定期代合計の画面表示用計算
- 定期券写真のアップロード
- 前月分定期申請Excelのアップロード
- 前月Excel添付時の入力項目制御
- `/generate` APIへのJSON送信
- 生成済みExcelのダウンロード開始

前月Excelが添付された場合、以下の項目のみを入力対象として残します。

- 提出日
- 上記理由が生じた年月日
- 利用開始日
- 定期券の写真

それ以外の項目は、前月Excelから読み取る前提となるため、画面上では非表示にします。

※/applyFilesCreater/src/main/resources/htmlに格納されていますが、サンプルのために
  同梱しているだけです。実体はS3オリジンに格納しています。

### request_carfare.html

`request_carfare.html` は、交通費精算書を作成するためのHTMLフォームです。

利用者が画面上で交通費明細を入力し、入力内容をJSON形式に変換して `/carfare_generate` APIへ送信します。  
APIから返却された署名付きダウンロードURLを使用して、生成された交通費精算書Excelをダウンロードします。

この画面では、定期申請書とは異なり、画像ファイルや前月Excelの添付は行いません。  
入力された交通費明細をもとに、Java側の `KotsuhiExcelService` が交通費精算書テンプレートへ値を転記します。

※/applyFilesCreater/src/main/resources/htmlに格納されていますが、サンプルのために
  同梱しているだけです。実体はS3オリジンに格納しています。

### Env.java

環境変数を管理するクラスです。

S3バケット名、テンプレートキー、一時アップロード先プレフィックスなどを定義します。

主な設定値は以下です。

| 項目 | 内容 |
|---|---|
| `REGION` | AWSリージョン |
| `BUCKET` | 使用するS3バケット名 |
| `TEMP_PREFIX` | 一時アップロードファイルの保存先プレフィックス |
| `TEMPLATE_KEY` | 定期申請書テンプレートのS3キー |
| `CARFARE_TEMPLATE_KEY` | 交通費精算書テンプレートのS3キー |
| `OUTPUT_PREFIX` | 生成済みExcelの出力先プレフィックス |
| `PASSES_PRICE_COL` | 定期代金額を書き込むExcel列 |

### LambdaHandler.java

API Gatewayから呼び出されるLambdaの入口クラスです。

`RequestStreamHandler` を実装し、HTTPメソッドとパスに応じて処理を振り分けます。

対応している主なAPIは以下です。

| パス | 内容 |
|---|---|
| `/upload-url` | S3アップロード用の署名付きPUT URLを発行 |
| `/generate` | 定期申請書Excelを生成 |
| `/carfare_generate` | 交通費精算書Excelを生成 |

`/generate` では、前月Excelの `file_key` が送信されている場合、`TeikiExcelReadService` に処理を渡して前月Excelの内容を読み取ります。

また、処理終了後には一時アップロードファイルをS3から削除します。

### S3Service.java

S3操作をまとめたサービスクラスです。

主な機能は以下です。

- S3オブジェクトをローカルファイルへダウンロード
- S3オブジェクトをバイト配列として取得
- ローカルファイルをS3へアップロード
- 署名付きPUT URLの発行
- 署名付きGET URLの発行
- S3オブジェクトの削除

LambdaHandlerやExcel生成サービスから利用されます。

### TeikiExcelService.java

定期申請書Excelを生成するサービスクラスです。

S3から取得したExcelテンプレートを読み込み、リクエストJSONの内容を所定のセルへ転記します。

主な処理内容は以下です。

- 提出日の転記
- 勤務地名の転記
- 勤務先住所の転記
- 申請者氏名の転記
- 届出理由のチェック反映
- 理由発生日の転記
- 定期経路明細の転記
- 片道総通勤時間の計算
- 定期券写真の貼り付け
- 数式の再計算
- 出力ファイル名の生成

定期券写真はS3から取得し、Excelの「定期購入履歴」シートへ貼り付けます。

### TeikiExcelReadService.java

前月分の定期申請Excelを読み取るためのサービスクラスです。

前月Excelが添付された場合、LambdaHandlerから呼び出されます。

主な役割は以下です。

- 前月Excelを読み込み
- 勤務地名を取得
- 勤務先住所を取得
- 申請者氏名を取得
- 届出理由を取得
- 定期経路明細を取得
- 今回入力された提出日・理由発生日・利用開始日・写真情報と結合
- `TeikiExcelService` が処理できるJSON形式へ変換

これにより、利用者は前月分と同じ定期経路を再入力せずに申請書を作成できます。

### KotsuhiExcelService.java

交通費精算書Excelを生成するサービスクラスです。

リクエストJSONの交通費明細情報を読み取り、交通費精算書テンプレートへ転記します。

主な処理内容は以下です。

- 申請年月の転記
- 社員Noの転記
- 申請者氏名の転記
- プロジェクト名の転記
- 交通費明細の転記
- 片道／往復の表示変換
- 交通費区分の表示変換
- 数式の再計算
- 出力ファイル名の生成

金額合計はExcelテンプレート側の数式で計算する前提のため、Java側では合計金額を直接セルに設定しません。

### ApiResponse.java

API Gatewayへ返却するレスポンスを作成するユーティリティクラスです。

主な役割は以下です。

- HTTPステータスコードの設定
- JSONレスポンス本文の作成
- CORSヘッダーの設定
- `OPTIONS` リクエストへの対応

## API仕様

### POST /upload-url

S3へファイルをアップロードするための署名付きPUT URLを発行します。

#### リクエスト例

```json
{
  "file_name": "pass.jpg",
  "content_type": "image/jpeg"
}
```

#### レスポンス例

```json
{
  "put_url": "https://...",
  "file_key": "tmp/20260501/xxxxxxxx.jpg",
  "expires_in": 600
}
```

### POST /generate

定期申請書Excelを生成します。

#### 通常入力時のリクエスト例

```json
{
  "request_id": "xxxxxxxx",
  "input_mode": "manual",
  "submitted_at": "2026-05-01",
  "work_location": {
    "name": "○○プロジェクト",
    "address": "東京都..."
  },
  "applicant": {
    "name": "山田 太郎"
  },
  "notification": {
    "reason_text": "継続",
    "reason_effective_date": "2026-05-01"
  },
  "commute": {
    "usage_start_date": "2026-05-01",
    "passes": [
      {
        "transportation": "JR",
        "section": "横浜～品川",
        "one_way_minutes": 30,
        "purchase_period_text": "1ヶ月",
        "amount_yen": 10000,
        "note": ""
      }
    ],
    "pass_photos": []
  }
}
```

#### 前月Excel利用時のリクエスト例

```json
{
  "request_id": "xxxxxxxx",
  "input_mode": "previous_excel",
  "submitted_at": "2026-05-01",
  "notification": {
    "reason_effective_date": "2026-05-01"
  },
  "commute": {
    "usage_start_date": "2026-05-01",
    "pass_photos": [],
    "previous_application_excel": {
      "file_key": "tmp/20260501/xxxxxxxx.xlsx",
      "file_name": "前月分.xlsx",
      "content_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    }
  }
}
```

#### レスポンス例

```json
{
  "download_url": "https://...",
  "file_key": "temp/20260501/xxxxxxxx.xlsx",
  "expires_in": 600
}
```

### POST /carfare_generate

交通費精算書Excelを生成します。

#### リクエスト例

```json
{
  "request_id": "xxxxxxxx",
  "submit_date": "2026-05",
  "applicant_no": "12345",
  "applicant_name": "山田 太郎",
  "applicant_project_name": "○○プロジェクト",
  "total_count": 1,
  "total_amount": 1000,
  "carfare_details": [
    {
      "carfare_date": "2026-05-01",
      "carfare_destination": "本社",
      "carfare_destination_line": "JR",
      "carfare_reason": "打合せ",
      "carfare_departure": "横浜",
      "carfare_arrival": "品川",
      "carfare_unit_price": 500,
      "carfare_one_way_or_round_trip": "round_trip",
      "carfare_section": "return_to_office",
      "carfare_amount": 1000
    }
  ]
}
```

#### レスポンス例

```json
{
  "download_url": "https://...",
  "file_key": "temp/20260501/xxxxxxxx.xlsx",
  "expires_in": 600
}
```

## S3キー構成

### 一時アップロードファイル

```text
tmp/yyyyMMdd/ランダム文字列.拡張子
```

例：

```text
tmp/20260501/abcd1234.xlsx
tmp/20260501/efgh5678.jpg
```

一時アップロードファイルは、処理終了後にJava側で削除します。

### 生成済みExcelファイル

```text
temp/yyyyMMdd/ランダム文字列.xlsx
```

例：

```text
temp/20260501/abcd1234.xlsx
```

生成済みExcelは、署名付きGET URLを利用してブラウザからダウンロードします。

## 環境変数

| 環境変数 | 必須 | 初期値 | 説明 |
|---|---|---|---|
| `AWS_REGION` | 任意 | `ap-northeast-1` | AWSリージョン |
| `BUCKET` | 必須 | 空文字 | S3バケット名 |
| `TEMP_PREFIX` | 任意 | `tmp` | 一時アップロード先 |
| `TEMPLATE_KEY` | 任意 | `template/定期申請書YYYY年度MM月(氏名).xlsx` | 定期申請書テンプレート |
| `CARFARE_TEMPLATE_KEY` | 任意 | `template/交通費精算書YYYY年度MM月(氏名).xlsx` | 交通費精算書テンプレート |
| `OUTPUT_PREFIX` | 任意 | `temp` | 生成済みExcel出力先 |
| `PASSES_PRICE_COL` | 任意 | `K` | 定期代金額列 |

## セキュリティ上の注意

### S3削除対象の制限

一時ファイル削除処理では、削除対象のキーが `TEMP_PREFIX` 配下であることを確認してから削除します。

これにより、以下のような重要ファイルを誤って削除しないようにしています。

- Excelテンプレート
- 生成済みExcel
- その他S3上の管理ファイル

### 署名付きURLの有効期限

アップロード用・ダウンロード用の署名付きURLには有効期限を設定しています。

現在の想定では以下です。

```text
600秒
```

つまり、発行から10分程度でURLは無効になります。

### ファイル本体の扱い

ファイル本体はAPI GatewayやLambdaのリクエスト本文には直接含めません。

画像やExcelは、ブラウザからS3へ直接アップロードします。  
LambdaにはS3上の `file_key` のみを渡します。

これにより、API GatewayやLambdaのリクエストサイズ制限を受けにくくなります。

## 補足事項

### 前月Excel添付時の扱い

前月Excelが添付された場合、HTMLフォーム上では一部の入力欄を非表示にします。

これは、前月Excelから以下の情報を取得するためです。

- 勤務地名
- 勤務先住所
- 申請者氏名
- 届出理由
- 定期経路明細

一方で、以下の情報は今回申請固有の情報として入力します。

- 提出日
- 上記理由が生じた年月日
- 利用開始日
- 定期券の写真

### 合計金額について

画面上では合計金額を表示しますが、最終的なExcel上の合計はテンプレート側の数式に任せています。

そのため、Java側では原則として合計金額セルへ直接値を書き込みません。

### Lambdaのローカルファイル

Lambda実行中は `/tmp` 配下に一時ファイルを保存します。

主な一時ファイルは以下です。

```text
/tmp/template.xlsx
/tmp/previous_application.xlsx
/tmp/output.xlsx
```

処理終了後、これらのローカル一時ファイルは `finally` で削除します。

## 今後の改善案

### 生成済みExcelの削除

現在、生成済みExcelはS3へ保存され、署名付きGET URLでダウンロードします。

より厳密に管理する場合は、以下の方法が考えられます。

- ブラウザでダウンロード完了後に削除APIを呼ぶ
- S3ライフサイクルルールで一定期間後に削除する
- Lambdaから直接ファイルを返却し、S3へ完成ファイルを残さない構成にする

実務上は、削除APIとS3ライフサイクルの併用が安全です。

### ファイル種別チェックの強化

現在は拡張子やContent-Typeをもとにファイルを扱います。

より安全にする場合は、以下のチェックを追加するとよいです。

- Excelファイルの拡張子チェック
- 画像ファイルのContent-Typeチェック
- ファイルサイズ上限チェック
- 想定外ファイルの場合のエラーレスポンス返却

### ログ出力の整理

現状は `System.out.println` によるログ出力が中心です。

運用を考える場合は、以下の観点でログを整理するとよいです。

- リクエストID
- 処理開始・終了
- S3キー
- エラー内容
- 削除処理の成否

### エラー内容のユーザー向け整理

内部エラーをそのまま画面に出すと、利用者には分かりにくい場合があります。

本番運用では、画面表示用メッセージとログ出力用メッセージを分けるとよいです。

## 開発メモ

### Excelテンプレートのセル位置

定期申請書テンプレートの主なセル位置は以下です。

| 項目 | セル |
|---|---|
| 提出日 | K3 |
| 勤務地名 | C5 |
| 勤務先住所 | C6 |
| 申請者氏名 | C7 |
| 届出理由 | M5 |
| 交通機関名 | B10〜B15 |
| 移動区間 | D10〜D15 |
| 所要時間 | H10〜H15 |
| 購入期間 | I10〜I15 |
| 定期代金額 | K10〜K15 |
| 備考 | N10〜N15 |
| 片道総通勤時間 | E16 |

交通費精算書テンプレートの主なセル位置は以下です。

| 項目 | セル |
|---|---|
| 提出日 | AD2 |
| 申請月 | E10 |
| 社員No | B5 |
| 申請者氏名 | H5 |
| プロジェクト名 | H7 |
| 月日 | C13〜 |
| 行き先 | E13〜 |
| 利用路線 | I13〜 |
| 理由 | M13〜 |
| 出発 | S13〜 |
| 到着 | V13〜 |
| 単価 | Y13〜 |
| 片道／往復 | AA13〜 |
| 交通費区分 | AC13〜 |

## 注意

Excelテンプレートのレイアウトを変更した場合は、Java側のセル指定も合わせて修正する必要があります。

特に以下の変更には注意してください。

- シート名の変更
- セル位置の変更
- 明細行の開始行・終了行の変更
- 定期代金額列の変更
- 写真貼り付け位置の変更
- 数式セルの変更

