# typical-reply-outlook

## 概要

定型文での返信を提供するOutlook向けアドオンです。

例えば、以下のメールを受信したとします。

```
件名:
  例の件について

本文:
  お疲れ様です。
  例の件について以下の案を考えてみたのですが、いかがでしょうか？
  http://...
```

このとき、以下のいずれかの手順ですぐに定型文で返信することができます。

* リボンの「Typical Reply」ギャラリーから、「いいね！」を選択
  !["リボン"](./Documents/Images/RibbonGallery.png "リボン")
* メールを右クリックして表示されるコンテキストメニューの「Typical Reply」ギャラリーから、「いいね！」を選択
  !["コンテキストメニュー"](./Documents/Images/ContextMenuGallery.png "コンテキストメニュー")

```
Subject:
  [[いいね！]]: Re: 例の件について

Body:
  いいね！
  
  > -----Original Message-----
  > お疲れ様です。
  > 例の件について以下の案を考えてみたのですが、いかがでしょうか？
  > http://...
```

これらの返信内容の設定はグループポリシーによってドメイン単位での一元管理が可能です。
個別の設定ファイルによるユーザー単位での設定変更も可能です。

## グループポリシーによる設定

policy配下のadmx、admlファイルをグループポリシーファイルを以下のパスに配置します。

* Active Directoryのグループポリシーに追加する場合
  * `C:\Windows\SYSVOL\domain\Policies\PolicyDefinitions`
* ローカルグループポリシーに追加する場合
  * `C:\Windows\PolicyDefinitions`

グループポリシーエディターを開き、 管理用テンプレート -> TypicalReply -> 既定値 -> TypicalReply設定 を開きます。
この設定を有効にし、テキストエリアに後述のJSONでの設定を記載します。

## 設定ファイルによる設定

設定ファイルは以下の箇所に存在します。

`%APPDATA%\TypicalReply\TypicalReplyConfig.json`

このファイルに後述のJSONでの設定を記載します。

実際に使用される設定は、グループポリシーの内容と設定ファイルによる内容がマージされた設定となります。

* GalleryLabelは設定ファイルのもの使用されます（指定していた場合）
* ButtonConfigListはグループポリシーの内容と設定ファイルの内容がマージされます
  * IDに重複がある場合、設定ファイルのものが使用されます

## インストーラーでのデフォルト設定ファイルのインストール

セットアップの存在するフォルダに `DefaultConfig\TypicalReplyConfig.json` を配置した状態でセットアップを実行することで、
セットアップ実行時に自動で`%APPDATA%\TypicalReply\TypicalReplyConfig.json` にファイルが配置されます。

## 設定項目

設定は以下のようなJSON形式で指定します。

```
{
    "Priority" : 1,
    "ConfigList": [
        {
            "Culture": "ja-JP",
            "GalleryLabel": "定型返信",
            "ButtonConfigList": [
                {
                    "Id": "Like",
                    "Label": "いいね！",
                    "SubjectPrefix": "[[いいね！]]:",
                    "Body": "いいね！",
                    "Recipients": ["all"],
                    "QuoteType": true,
                    "AllowedDomains": [
                        "*"
                    ]
                },
                {
                    "Id": "OK",
                    "Label": "了解",
                    "SubjectPrefix": "[[了解]]:",
                    "Body": "了解しました。",
                    "Recipients": ["all"],
                    "QuoteType": true,
                    "AllowedDomains": [
                        "*"
                    ]
                }
            ]
        },
        {
            "Culture": "en-US",
            "GalleryLabel": "Typical Reply",
            "ButtonConfigList": [
                {
                    "Id": "Like",
                    "Label": "Like!",
                    "SubjectPrefix": "[[Like!]]:",
                    "Body": "Like!",
                    "Recipients": ["all"],
                    "QuoteType": true,
                    "AllowedDomains": [
                        "*"
                    ]
                },
                {
                    "Id": "OK",
                    "Label": "OK",
                    "SubjectPrefix": "[[OK]]:",
                    "Body": "OK.",
                    "Recipients": ["all"],
                    "QuoteType": true,
                    "AllowedDomains": [
                        "*"
                    ]
                }
            ]
        }
    ]
}
```

TypiclReplyConfig: 設定のルート

| 設定名     | 型             | 必須 | 省略時のデフォルト | 概要                                                                                                                         |
| ---------- | -------------- | ---- | ------------------ | ---------------------------------------------------------------------------------------------------------------------------- |
| Priority   | 数値           | no   | -1                 | レジストリと設定ファイルのどちらの設定を使うかの優先度。<br>値が大きい方が使用される。<br>値が同じ場合は設定がマージされる。 |
| ConfigList | Configのリスト | yes  | -                  | 各言語ごとの設定（Config）のリスト                                                                                           |


Config: 各言語ごとの設定

| 設定名           | 型                   | 必須 | 省略時のデフォルト | 概要                                                                                                                                                                     | 例                |
| ---------------- | -------------------- | ---- | ------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | ----------------- |
| Culture          | 文字列               | no   | null               | 対象となるカルチャ。<br>ロケールなしの言語のみを指定することも可能です。<br>現在のカルチャにマッチするCultureがない場合、Cultureの値に関わらず先頭のConfigを使用します。 | `"ja-JP"`、`"ja"` |
| GalleryLabel     | 文字列               | no   | "Typical Reply"    | リボンやコンテキストメニューに表示される本機能のラベル                                                                                                                   | `"定型返信"`      |
| ButtonConfigList | ButtonConfigのリスト | yes  | -                  | 定型返信ボタン設定のリスト                                                                                                                                               | -                 |

ButtonConfig: 定型返信ボタン設定。返信内容や返信先等の設定を行う。
| 設定名         | 型             | 必須 | 省略時のデフォルト   | 概要                                                                                                                                                                                        | 例                                        |
| -------------- | -------------- | ---- | -------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------- |
| Id             | 文字列         | yes  | -                    | ボタンのID。ButtonConfigList内で重複不可。                                                                                                                                                  | `"LikeId"`                                |
| Label          | 文字列         | yes  | -                    | ボタンに表示されるラベル                                                                                                                                                                    | `"いいね！"`                              |
| SubjectPrefix  | 文字列         | no   | null                 | 件名の先頭に挿入する文言                                                                                                                                                                    | `"[[いいね]]"`                            |
| Subject        | 文字列         | no   | 返信のデフォルト件名 | 件名                                                                                                                                                                                        | `"報告"`                                  |
| Body           | 文字列         | no   | null                 | 本文                                                                                                                                                                                        | `"いいね"`                                |
| Recipients     | 文字列のリスト | no   | 送信先なし           | 送信先。<br>`["blank"]`: 送信先なし<br> `["all"]`: 全員に返信<br>`["sender"]`: 送信者にだけ返信<br>その他の文字列リスト: 指定のアドレスに返信                                               | `["test@test.co.jp", "test2@test.co.jp"]` |
| QuoteType      | boolean        | no   | false                | 元の文言を引用するかどうか。 <br> `true`: 引用する<br>`false`: 引用しない                                                                                                                   | `true`                                    |
| AllowedDomains | 文字列のリスト | no   | 全て許可             | 送信を許可するドメインリスト。このドメイン以外が含まれている場合、返信用メールの作成、送信は行わない。<br>`["*"]`: 全て許可する<br>その他の文字列リスト: 指定したドメインのみ送信を許可する | `["test.co.jp", "test2.co.jp"]`           |
| ForwardType    | 文字列         | no   | 添付しない           | 元のメールを添付するかどうか。<br>`attachment`: 添付する                                                                                                                                    | `attachment`                              |

## 新しい設定追加の例

「最高！」というボタンを追加する方法を考えます。

設定ファイル（`%APPDATA%\TypicalReply\TypicalReplyConfig.json`）を編集します。

現在の設定が以下のようになっているとします。

```
{
    "ConfigList": [
        {
            "Culture": "ja-JP",
            "GalleryLabel": "定型返信",
            "ButtonConfigList": [
                {
                    "Id": "Like",
                    "Label": "いいね！",
                    "SubjectPrefix": "[[いいね！]]:",
                    "Body": "いいね！",
                    "Recipients": ["all"],
                    "QuoteType": true,
                    "AllowedDomains": [
                        "*"
                    ]
                }
            ]
        }
    ]
}
```

ButtonConfigListにButtonConfigを追加します。

`Id`は`Awesome`とし、`Label`は`最高！`とします。

```
{
    "Id": "Awesome",
    "Label": "最高！"
}
```

元のメッセージに対して返信するので、元の件名は残して、件名に対してリアクションのメッセージを追加します。
そのために、`Subject`は空にして元の件名が残るようにし、`SubjectPrefix`で件名の先頭にメッセージを追加します。

```
{
    "Id": "Awesome",
    "Label": "最高！",
    "SubjectPrefix": "[[最高！]]:"
}
```

同様に、元のメッセージに対して返信するので、元の本文は残して（引用状態にして）、本文にメッセージを追加します。
そのために、`Body`にメッセージを指定し、`QuoteType`に`true`を指定します。

```
{
    "Id": "Awesome",
    "Label": "最高！",
    "SubjectPrefix": "[[最高！]]:",
    "Body": "最高！",
    "QuoteType": true
}
```

このボタンでは、送信者にのみ返信することとします。
そのために、`Recipients`に`["sender"]`を指定します。

```
{
    "Id": "Awesome",
    "Label": "最高！",
    "SubjectPrefix": "[[最高！]]:",
    "Body": "最高！",
    "QuoteType": true,
    "Recipients": ["sender"]
}
```

また、送信先のドメインは自身が所属している`test.co.jp`のみに限定することとします。
そのために、`AllowedDomains`に`["all"]`を指定します。

```
{
    "Id": "Awesome",
    "Label": "最高！",
    "SubjectPrefix": "[[最高！]]:",
    "Body": "最高！",
    "QuoteType": true,
    "Recipients": ["sender"],
    "AllowedDomains": ["test.co.jp"]
}
```

元のメッセージの添付は不要とします。
そのため、`ForwardType`は指定しません。

以上で作成した設定をButtonConfigListに追加します。

```
{
    "ConfigList": [
        {
            "Culture": "ja-JP",
            "GalleryLabel": "定型返信",
            "ButtonConfigList": [
                {
                    "Id": "Like",
                    "Label": "いいね！",
                    "SubjectPrefix": "[[いいね！]]:",
                    "Body": "いいね！",
                    "Recipients": ["all"],
                    "QuoteType": true,
                    "AllowedDomains": [
                        "*"
                    ]
                },
                {
                    "Id": "Awesome",
                    "Label": "最高！",
                    "SubjectPrefix": "[[最高！]]:",
                    "Body": "最高！",
                    "QuoteType": true,
                    "Recipients": ["sender"],
                    "AllowedDomains": ["test.co.jp"]
                }
            ]
        }
    ]
}
```

これで、定型返信のボタンの中に、「最高！」ボタンが追加されます。

!["「最高！」ボタン"](./Documents/Images/Awesome.png "「最高！」ボタン")
