# IyakuhinDataSuite
医薬品データベースの作成

## 概要
複数のソースから医薬品データを収集し、YJコード(個別医薬品コード)をキーとして統合するツール群です。

## 主要スクリプト

### 1. データ収集
- **medhot_get.ps1**: MEDHOTから医薬品コードデータをダウンロード
- **pmda_yjcode_size_extract.ps1**: PMDA XMLファイルからYJコードと物理サイズ情報を抽出

### 2. データ処理
- **data_process.ps1**: CSVファイルをJSON形式に変換

### 3. データ統合
- **json_merge_yjcode_with_packaging.ps1**: 全データを統合し、包装単位コードを含むJSONを生成
  - 統合キー: 個別医薬品コード(YJCode)
  - 包装情報: 薬価基準収載医薬品コードを参照して取得
  - 出力場所: `data/integrated/`
  - 含まれる情報:
    - HOT9データ(基本医薬品情報)
    - PMDAサイズデータ(錠剤・カプセルの物理サイズ)
    - 包装単位コード(調剤包装単位コード、販売包装単位コード、元梱包装単位コード)

## 使用方法

```powershell
# 1. MEDHOTデータのダウンロード
.\medhot_get.ps1

# 2. PMDAデータの抽出
.\pmda_yjcode_size_extract.ps1

# 3. CSVをJSONに変換
.\data_process.ps1

# 4. データ統合(包装単位コード含む)
.\json_merge_yjcode_with_packaging.ps1
```

## 出力データ構造

```json
{
  "Metadata": {
    "GeneratedAt": "2025-10-05 23:53:33",
    "Description": "個別医薬品コード(YJCode)で統合された医薬品データ（包装単位コード含む）",
    "Statistics": {
      "TotalProducts": 26671,
      "TotalYJCodes": 26878,
      "YJCodesWithPackaging": 25139
    }
  },
  "Products": [
    {
      "ProductName": "製品名",
      "Variants": [
        {
          "YJCode": "1234567X1234",
          "HOT9": { /* 基本情報 */ },
          "PMDA": { /* サイズ情報 */ },
          "PackagingInfo": [
            {
              "調剤包装単位コード": "04987...",
              "販売包装単位コード": "14987...",
              "元梱包装単位コード": "24987...",
              "薬価コード": "1234567X1234"
            }
          ]
        }
      ]
    }
  ]
}
```

## データソース
- **MEDHOT**: https://medhot.medd.jp/view_download
- **HOT9**: MEDISマスターデータ
- **PMDA**: 医薬品医療機器総合機構XMLデータ

## 注意事項
- YJコード(個別医薬品コード)を統合キーとして使用
- 薬価コードや薬価基準収載医薬品コードは統合キーとしては使用せず、包装情報の取得にのみ使用
- 包装単位コードの付与率: 約93.5%
