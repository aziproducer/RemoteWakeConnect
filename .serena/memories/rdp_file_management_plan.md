# RDPファイル管理機能実装計画

## 現在の実装状況
- rdp_filesフォルダはcreate-release.ps1で言及されているが、実際のコードでは未実装
- ConnectionHistoryServiceでrdp_filesフォルダパスが定義されているが、使用されていない
- 外部RDPファイル読み込み機能は実装済み（BrowseRdpButton_Click）
- 直打ち接続機能は実装済み（ConnectButton_Click）

## 実装内容
1. rdp_filesフォルダの作成と管理
   - アプリケーション起動時にrdp_filesフォルダを自動作成
   - 外部RDPファイル読み込み時にrdp_filesフォルダにコピー

2. RDPファイル管理機能
   - 外部から読み込んだRDPファイルをrdp_filesフォルダにコピー
   - ファイル名は「{ComputerName/IP}_{DateTime}.rdp」形式で保存
   - 直打ち接続時にRDPファイルを自動生成

3. 履歴管理の更新
   - ConnectionHistoryServiceでrdp_filesフォルダのファイルパスを管理
   - 履歴からの接続時はrdp_filesフォルダのファイルを参照