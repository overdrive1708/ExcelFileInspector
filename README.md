[日本語](README.md)

<h1 align="center">
    <a href="https://github.com/overdrive1708/ClipboardNumberToName">
        <img alt="ClipboardNumberToName" src="docs/images/AppIconReadme.png" width="50" height="50">
    </a><br>
    ExcelFileInspector
</h1>

<h2 align="center">
    エクセルファイルをチェックするアプリケーション
</h2>

<div align="center">
    <img alt="csharp" src="https://img.shields.io/badge/csharp-blue.svg?style=plastic&logo=csharp">
    <img alt="dotnet" src="https://img.shields.io/badge/.NET-blue.svg?style=plastic&logo=dotnet">
    <img alt="license" src="https://img.shields.io/github/license/overdrive1708/ExcelFileInspector?style=plastic">
    <br>
    <img alt="repo size" src="https://img.shields.io/github/repo-size/overdrive1708/ExcelFileInspector?style=plastic&logo=github">
    <img alt="release" src="https://img.shields.io/github/release/overdrive1708/ExcelFileInspector?style=plastic&logo=github">
    <img alt="download" src="https://img.shields.io/github/downloads/overdrive1708/ExcelFileInspector/total?style=plastic&logo=github&color=brightgreen">
    <img alt="open issues" src="https://img.shields.io/github/issues-raw/overdrive1708/ExcelFileInspector?style=plastic&logo=github&color=brightgreen">
    <img alt="closed issues" src="https://img.shields.io/github/issues-closed-raw/overdrive1708/ExcelFileInspector?style=plastic&logo=github&color=brightgreen">
</div>

## ダウンロード方法

[GitHubのReleases](https://github.com/overdrive1708/ExcelFileInspector/releases)にあるLatestのAssetsよりExcelFileInspector_vx.x.x.zipをダウンロードしてください｡

## 初回セットアップ

1. ｢ExcelFileInspector.exe｣を一回起動して終了してください｡｢InspectionSettings.json｣が生成されます｡
1. [設定例](docs/設定例InspectionSettings.json)を参考に｢InspectionSettings.json｣を記載してください｡

｢InspectionSettings.json｣で設定する設定項目は以下のとおりです｡
| 設定項目 | 設定内容 |
| --- | --- |
| PresetName | プリセットを識別するための名前を設定します｡ |
| TargetFileKeyword | 検査ファイルの一覧に登録する際のキーワードを設定します｡<br>検査ファイルに検査対象フォルダもしくは検査対象ファイルをドラッグ&ドロップした際に､設定したキーワードを含むファイルのみを検査対象とします｡ |
| InspectionMethods | Conditionによって設定値が異なるため､後述する説明を参照してください｡ |

｢InspectionSettings.json｣で設定するConditionは以下をサポートしています｡
| Condition | 設定時の挙動 |
|---|---|
| Equal | SheetNameで設定したシート名の､Cellで設定したセルが､Valueで設定した値である場合にNGとします｡ |
| NotEqual | SheetNameで設定したシート名の､Cellで設定したセルが､Valueで設定した値以外である場合にNGとします｡ |
| Empty | SheetNameで設定したシート名の､Cellで設定したセルが､空である場合にNGとします｡ |
| NotEmpty | SheetNameで設定したシート名の､Cellで設定したセルが､空ではない場合にNGとします｡ |

## 使用方法

1. ｢ExcelFileInspector.exe｣を起動してください｡
1. 検査設定をプリセットから選択してください｡
1. 検査ファイルに検査対象フォルダもしくは検査対象ファイルをドラッグ&ドロップしてください｡
1. ｢検査実施｣ボタンを押下してください｡

## 開発環境

- Microsoft Visual Studio Community 2022

## 使用しているライブラリ

詳細は[NOTICE.md](NOTICE.md)を参照してください｡

## ライセンス

このプロジェクトはMITライセンスです。  
詳細は [LICENSE](LICENSE) を参照してください。

## 作者

[overdrive1708](https://github.com/overdrive1708)
