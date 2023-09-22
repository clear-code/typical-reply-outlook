## 設定ファイルテスト手順

### 各設定が正しく動くことの確認

Typical Replyに関するレジストリは空の状態で、設定ファイルによる動作確認を実施する。

設定ファイルは`%APPDATA%\TypicalReply\TypicalReplyConfig.json`である。
ログファイルは`%APPDATA%\TypicalReply\TypicalReplyConfig.log`である。

ConfigForTest/TypicalReplyConfig.jsonが言語のテスト用の設定である。
これを`%APPDATA%\TypicalReply\TypicalReplyConfig.json`に配置することでテストを実施する。

Windowsの表示言語はWindowsの「設定->言語->Windowsの表示言語」から表示言語を変更する。

* Windowsの表示言語が日本語の時、Cultureがjaの設定が使われること
  * ギャラリー名が「テストギャラリー」であること
  * ボタン名が「テストボタン」であること
  * 返信時の件名が[[日本語]]であること
* Windowsの表示言語が英語の時、Cultureがenの設定が使われること
  * ギャラリー名が「Test Gallery」であること
  * ボタン名が「Test Button」であること
  * 返信時の件名が[[English]]であること
* Windowsの表示言語が中国語簡体字の時、Cultureがen-CNの設定が使われること
  * ギャラリー名が「测试画廊」であること
  * ボタン名が「测试按钮」であること
  * 返信時の件名が[[简体字]]であること
* Windowsの表示言語が中国語繁体字の時、Cultureがen-TWの設定が使われること
  * ギャラリー名が「測試畫廊」であること
  * ボタン名が「測試按鈕」であること
  * 返信時の件名が[[繁體中文]]であること
