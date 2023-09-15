# typical-reply-outlook

## 概要

定型文での返信を提供するOutlook向けアドオンです。

例えば、以下のメールを受信したとします。

```
Subject:
  How about this?

Body:
  Hi.
  I've wrote a plan. How about this?
  http://...
```

このとき、以下のいずれかの手順ですぐに定型文で返信することができます。

* リボンの「Typical Reply」ギャラリーから、「いいね」を選択
* メールを右クリックして表示されるコンテキストメニューの「Typical Reply」ギャラリーから、「いいね」を選択します。

```
Subject:
  [[Like!]]: Re: How about this?

Body:
  Like!
  
  > Hi.
  > I've wrote a plan. How about this?
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

設定ファイルは以下の箇所に保存されています。

`%APPDATA%\TypicalReply\TypicalReplyConfig.json`

実際に使用される設定は、グループポリシーの内容と設定ファイルによる内容がマージされた設定となります。

* GalleryLabelは設定ファイルのもの使用されます（指定していた場合）
* ButtonConfigListはグループポリシーの内容と設定ファイルの内容がマージされます
  * IDに重複がある場合、設定ファイルのものが使用されます

## 新しい定型返信設定を追加するには

設定は以下のようなJSON形式で指定します。

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
                    "SubjectPrefix": "[[いいね！]]",
                    "Body": "いいね！",
                    "Recipients": ["*"],
                    "QuoteType": true,
                    "AllowedDomains": [
                        "*"
                    ]
                },
                {
                    "Id": "OK",
                    "Label": "了解",
                    "SubjectPrefix": "[[了解]]",
                    "Body": "了解しました。",
                    "Recipients": ["*"],
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
                    "SubjectPrefix": "[[Like!]]",
                    "Body": "Like!",
                    "Recipients": ["*"],
                    "QuoteType": true,
                    "AllowedDomains": [
                        "*"
                    ]
                },
                {
                    "Id": "OK",
                    "Label": "OK",
                    "SubjectPrefix": "[[OK]]",
                    "Body": "OK.",
                    "Recipients": ["*"],
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
|設定名|型|必須|省略時のデフォルト|概要|
|--|--|--|--|--|--|
|ConfigList|Configのリスト|yes|-|各言語ごとの設定（Config）のリスト|


Config: 各言語ごとの設定
|設定名|型|必須|省略時のデフォルト|概要|例|
|--|--|--|--|--|--|
|Culture|文字列|no|null|対象となるカルチャ。ロケールなし言語のみを指定することも可能です。現在のカルチャにマッチするCultureがない場合、Cultureの値に関わらず先頭のConfigを使用します。|`"ja-JP"`、`"ja"`|
|GalleryLabel|文字列|no|"Typical Reply"|リボンやコンテキストメニューに表示される本機能のラベル|`"定型返信"`|
|ButtonConfigList|ButtonConfigのリスト|yes|-|定型返信ボタン設定のリスト|-|

