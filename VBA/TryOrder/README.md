# TryOrder (Selenium Basic Wrapper Class)

TryOrderは、VBAにおけるSelenium Basicの操作を安定化させるためのラッパークラスモジュールです。ブラウザ操作の不安定要素（DOMの読み込み待ちや要素の未検出）をリトライ処理によって解決し、堅牢なスクレイピングの実装を支援します。

## 概要

本クラスは、WebDriverの各操作をプロパティベースで制御します。要素が見つからない場合に即座にエラーを出すのではなく、指定したタイムアウト時間内でのループ試行（ポーリング）を行うことが最大の特徴です。

## 主な機能

* **自動リトライ（待機）処理**: 要素のクリックやテキスト入力時、要素が表示されるまで指定秒数リトライを繰り返します。
* **多彩なセレクター対応**: ID, Name, CSS Selector, XPath, LinkTextによる要素特定が可能です。
* **ウィンドウ・アラート管理**: ウィンドウタイトルによる動的な切り替えや、出現タイミングの不安定なアラートの捕捉が可能です。
* **スクロール・アクション対応**: `scrollIntoView`による要素へのスクロールや、マウスムーブ（MoveToElement）をサポートしています。

---

## 構成・設定

### 検索モード (findMode)
`SelectorFind`プロパティに指定します。
* `byId`
* `byName`
* `byCss`
* `byXPath`
* `byLinkText`

### アクション型 (methodType)
`DoAction`または`GetResult`メソッドの引数として使用し、要素に対する操作を決定します。

| 定数 | 概要 |
| :--- | :--- |
| mClick | 要素をクリックする |
| mSendKeys | 値を入力する（InputTextプロパティを使用） |
| mClear | 入力内容をクリアする |
| mInnerText | 要素内のテキストを取得する |
| mGetValue | Value属性を取得する |
| mScrollIntoView | 要素が画面中央に来るようスクロールする |
| mAsSelect | プルダウンメニューを選択する |

---

## 実装例

### 1. 基本操作（要素の待機とクリック）
要素が表示されるまで最大10秒間待機してからクリックを実行します。

```vba
Dim order As New TryOrder
Set order.WebDriver = driver ' 既存のWebDriverインスタンスをセット

With order
    .Target = "submit_button_id"
    .SelectorFind = findMode.byId
    .TimeoutSec = 10
    .DoAction methodType.mClick
End With
```

### 2. 要素からの情報取得
要素のテキスト配列を取得します。

```vba
Dim results As Variant
With order
    .Target = ".item-list-names"
    .SelectorFind = findMode.byCss
    results = .GetResult(methodType.mInnerText)
End With
```

### 3. ウィンドウの切り替え
特定のタイトルが含まれるタブを探してアクティブにします。

```vba
With order
    .InputText = "対象のページタイトル"
    .TimeoutSec = 5
    If .FindWindowByTitle() Then
        Debug.Print "ウィンドウ切り替えに成功しました"
    End If
End With
```



## 依存関係および注意点

1. **Selenium Type Library** 参照設定に「Selenium Type Library」が必要です。
2. **ErrorLogクラス** 本コードはErrorLogという名前の独自ログ管理クラスを参照しています。環境に合わせて該当箇所をコメントアウト、または修正してください。
3. **インデックス** IndexNumberのデフォルト値は1に設定されています（SeleniumのItemインデックスに準拠）。

## その他
drft.clsはウィンドウタイトル指定をSwitchToWindowの繰り返しをせずショートカットしてフォーカスする為の考えの試作


## ライセンス
このプロジェクトは **[MITライセンス](https://opensource.org/licenses/mit-license.php)** の下で公開されています。
