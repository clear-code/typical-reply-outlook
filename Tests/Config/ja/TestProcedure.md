## 設定ファイルテスト手順

### 各設定が正しく動くことの確認

Typical Replyに関するレジストリは空の状態で、設定ファイルによる動作確認を実施する。

設定ファイルは`%APPDATA%\TypicalReply\TypicalReplyConfig.json`である。
ログファイルは`%APPDATA%\TypicalReply\TypicalReplyConfig.log`である。

ConfigForTest/TypicalReplyConfig.jsonが各ボタンのテスト用の設定である。
これを`%APPDATA%\TypicalReply\TypicalReplyConfig.json`に配置することでテストを実施する。

* 設定ファイルが存在しない場合、TypicalReplyが表示されないこと（エラーになること）
  * ファイルの読み込みに失敗した旨がログ出力されること
* 設定ファイルの記載に誤りがある場合、TypicalReplyが表示されないこと
  * ファイルの読み込みに失敗した旨がログ出力されること
* Config.Culture
  * 言語と地域（ja-JP）を指定する場合
    * 現在のカルチャ一致するカルチャが存在する場合、そのカルチャの設定が使用されること
    * 一致するカルチャが存在しない場合、先頭の設定が使用されること
    * 同じカルチャの言語のみ（ja）を指定した設定が存在する場合
      * 言語と地域の設定が優先されること
    * 現在のカルチャに一致する設定がない場合、先頭の設定が使用されること
  * 言語のみ（ja）を指定する場合
    * 現在のカルチャの言語と一致する言語が存在する場合、その言語の設定が使用されること
    * 現在のカルチャに一致する設定がない場合、先頭の設定が使用されること
  * 指定しない場合
    * 他に現在のカルチャに一致する設定がない場合、先頭の設定が使用されること
* Config.GroupLabel
  * 指定しない場合、 "Typical Reply"になること
  * 指定した場合、指定した値になること
* Config.ButtonConfigList
  * 指定しない場合、ボタンが何も表示されないこと
  * 指定した場合、ボタンが表示されること
    * 複数指定した場合は複数表示されること

* ButtonConfig.Id
  * 指定しない場合、このボタンが表示されないこと
  * Labelと共に指定した場合、このボタンが表示されること
* ButtonConfig.Label
  * 指定しない場合、このボタンが表示されないこと
  * Idと共に指定した場合、このボタンが表示されること
* ButtonConfig.SubjectPrefix
  * 指定しない場合、返信の件名の先頭に何も挿入されないこと
  * 指定した場合、返信の件名の先頭に指定した文字列が挿入されること
* ButtonConfig.Subject
  * 指定しない場合、デフォルトの件名が表示されること
  * 指定した場合、件名がこの文字列になること
* ButtonConfig.Body
  * 指定しない場合、本文がないこと
  * 指定した場合、本文がこの文字列になること
* ButtonConfig.Recipients
  * **配列型である点に注意**
  * 指定しない場合、返信先が存在しないこと
  * `["blank"]`を指定した場合、全員に返信されること
  * `["sender"]`を指定した場合、送信者のみに返信されること
  * `["all"]`を指定した場合、全員に返信されること
  * `["test@test.co.jp", "test@test2.co.jp"]`を指定した場合、この2つに返信されること
* ButtonConfig.QuoteType
  * 指定しない場合、返信の本文に引用がないこと
  * falseを指定した場合、返信の本文に引用がないこと
  * trueを指定した場合、返信の本文に引用があること
* ButtonConfig.AllowedDomains
  * **配列型である点に注意**
  * 指定しない場合、全てのドメインが許可されること
  * `["*"]`を指定した場合、全てのドメインが許可されること
  * `["test.co.jp", "test.co.jp"]`を指定した場合
    * `test.co.jp`と`test2.co.jp`を含まないドメインがあった場合、返信されないこと
      * ※返信のための新規メールも何も作成されないこと
    * `test.co.jp`と`test.co.jp`ドメインのみだった場合、返信されること
* Button.ForwardType
  * 指定しない場合、元のメッセージが添付されないこと
  * "attachment"を指定した場合、元のメッセージが添付されること
  * **"inlune"は現在未対応**

### レジストリからのデータの読み込み

#### 準備

TypicalReply for Outlookのpolicy配下のadmx、admlファイルをC:\Windows\PolicyDefinitionsにコピーする。
gpedit.mscからグループポリシーエディターを開き、「管理用テンプレート->TypicalReply->既定値->TypicalReply設定」を開く
この設定のマルチテキストボックスに、設定ファイル（`%APPDATA%\TypicalReply\TypicalReplyConfig.json`）と同様の内容を記載する。
その後、設定ファイルを削除する。

#### テスト

* Typical Replyの設定が正しく読み込まれ、グループポリシーに設定した内容でボタンが表示されていること
* 「管理用テンプレート->TypicalReply->既定値->TypicalReply設定」の内容を変更し、Outlookを再起動すると、変更内容が反映されていること
  * 例えば、ラベルを変更する

### レジストリと設定ファイルの設定内容のマージ

#### 準備

「レジストリからのデータの読み込み」の準備を実施し、レジストリ設定を作成する。
ただし、この時設定ファイル（`%APPDATA%\TypicalReply\TypicalReplyConfig.json`）は削除せず、レジストリの内容と同じ設定をしておく。
また、TypicalReplyConfig.Priorityは同じ値にする（共に省略でも可）。

#### テスト

* Config.GroupLabel
  * 設定ファイルの値を変更すると、設定ファイルの内容が優先されること
* ButtonConfig
  * 設定ファイルの設定とレジストリの設定で、IDが同じものについて、設定ファイルのラベルを変更すると、設定ファイルの内容で上書きされること
  * 設定ファイルにIDが重複しないボタン要素を追加すると、ButtonConfigListに追加されること

### レジストリと設定ファイルの優先度の設定

#### 準備

「レジストリと設定ファイルの設定内容のマージ」テストでテスト時に変更した、設定ファイルとレジストリ設定を使用する。
（つまり、Config.GroupLabelとConfig.ButtonConfigListに差異がある状態であること。）

#### テスト

* レジストリのTypicalReplyConfig.Priorityの値を1、設定ファイルのTypicalReplyConfig.Priorityの値を2にする。
  * 設定ファイルの設定が使用されること
* レジストリのTypicalReplyConfig.Priorityの値を2、設定ファイルのTypicalReplyConfig.Priorityの値を1にする。
  * レジストリの設定が使用されること