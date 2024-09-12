## 準備

* Visual Studio 2019をインストール
* Microsoft Office Developer ToolsをVisual Studioインストーラーからインストール
  * 参考: * https://learn.microsoft.com/en-us/visualstudio/vsto/walkthrough-creating-your-first-vsto-add-in-for-outlook?view=vs-2022&tabs=csharp
* Inno Setupをインストール
* Inno Setupのunofficialな中国語簡体字と繁体字のメッセージリソースをインストールする
  * Inno Setupのインストールフォルダー\Languages配下に展開すること
* Visual Studio起動時Nugetライブラリの復元を求められた場合は復元する
* VSTO開発用のダミー証明書を生成
  * ソリューションエクスプローラでFlexConfirmMailを右クリックします。
  * Properties > Signing （プロパティ > 署名）を順に選択します。
  * 証明書欄に何か妥当な証明書の情報が表示されている場合、ダミー証明書はすでに生成済みです。
    証明書の情報が表示されていない場合、ダミー証明書を作成します。
      1. Create Test Certificate（テスト証明書の作成）をクリックします。
      2. テスト証明書の作成画面が開かれるので、何も入力せず「OK」ボタンを押して操作を確定します（（パスワードは不要です）。
         * `アクセスが拒否されました。（HRESULTからの例外: 0.x80070005(E_ACCESSDENIED)）` のようなメッセージが表示されて証明書の作成に失敗する場合、以下の手順で代替可能です。
           1. PowerShellを起動し、`New-SelfSignedCertificate -Subject "CN=FlexConfirmMailSelfSign" -Type CodeSigningCert -CertStoreLocation "Cert:\CurrentUser\My" -NotAfter (Get-Date).AddYears(10)` を実行します。
           2. Visual Studioのウィンドウに戻り、Choose from Store（ストアから選択）をクリックします。
           3. 証明書データベースのクライアント証明書のうち1つが表示されます。
             「FlexConfirmMailSelfSign」が表示されている場合、「OK」ボタンをクリックして選択します。
             そうでない場合、「その他」をクリックしてクライアント証明書の一覧から「FlexConfirmMailSelfSign」を選択し、「OK」ボタンをクリックして選択します
  * このダミー証明書を登録したことにより発生したcsprojファイルの差分はコミットしないようにしてください。
* 開発用のOfficeアプリを用意します（任意）
  * コンパイルのみであれば必要ありませんが、用意すると開発上は便利です。

## デバッグ

ローカルにOfficeアプリがインストールされている場合、Visual Studioからデバッグを実行すると、ローカルのOfficeがアドインが有効な状態で起動します。
こうすることで、ブレークポイントを作成するなど、Visual Studioを用いた通常のデバッグが可能となります。

## インストーラー

### ビルド

#### テスト版

準備手順のVSTO開発用のダミー証明書を作成した状態でmake.batを実行することで、テスト用のインストーラーを作成可能です。
以下の成果物がdestに作成されます。

* TypicalReplySetup-x.x.x.x.exe
  * インストーラー
* TypicalReplyOutlook-x.x.x.x-with-Default-Config.zip
  * インストーラーと既定の設定をまとめたパッケージ
* TypicalReplyADMX.zip
  * グループポリシーのテンプレートをまとめたパッケージ

#### 製品版

製品版のインストーラーを作成する場合、手元に対応する証明書が必要です。

TypicalReply.csprojを準備手順のダミー証明書の作成をしていない状態にし、make_signed.batを実行することで、製品版向けの署名ありのインストーラーを作成可能です。

TypicalReply.csprojをダミー証明書の作成をしていない状態とするには、例えば以下のコマンドを実行します。

```
git checkout TypicalReply.csproj
```

以下の成果物がdestに作成されます。

* TypicalReplySetup-x.x.x.x.exe
  * インストーラー
* TypicalReplyOutlook-x.x.x.x-with-Default-Config.zip
  * インストーラーと既定の設定をまとめたパッケージ
* TypicalReplyADMX.zip
  * グループポリシーのテンプレートをまとめたパッケージ

### 既定値の設定

TypicalReplyOutlook-x.x.x.x-with-Default-Config.zipにはインストーラーの他に`DefaultConfig\TypicalReplyConfig.json`が含まれています。
インストーラーはこのファイルを既定の設定として`%APPDATA%\TypicalConfig\TypicalReplyConfig.json`に配置します。

このファイルはリポジトリの`DefaultConfig\TypicalReplyConfig.json`をコピーしているため、既定値を変更する場合はこちらのファイルを変更します。