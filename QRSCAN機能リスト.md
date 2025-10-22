# ItemPicking QRスキャン機能リスト

## 📱 概要

ItemPicking システムは、**QR-Scanner 1.4.2** ライブラリを使用したWebカメラベースのQRコードスキャン機能を実装しています。

---

## 🎯 主要機能一覧

### 1. QRスキャンライブラリ

#### 使用ライブラリ
- **QR-Scanner v1.4.2**
- CDN: `https://unpkg.com/qr-scanner@1.4.2/qr-scanner.umd.min.js`
- タイプ: UMD (Universal Module Definition)

#### 実装場所
```html
<!-- ItemPicking.html 12行目 -->
<script src="https://unpkg.com/qr-scanner@1.4.2/qr-scanner.umd.min.js"></script>
```

---

## 🎨 UIコンポーネント

### 2. QRスキャンセクション (`qr-section-inline`)

#### 配置
- 品目情報ヘッダーの右側に横並び表示
- レスポンシブデザイン対応（画面幅260px以下で縦並びに切り替え）

#### 構成要素
1. **ステータスメッセージエリア** (`qrStatusMessage`)
2. **操作ボタングループ** (`qr-controls`)
3. **ビデオプレビューコンテナ** (`videoContainer`)
4. **スキャン結果表示エリア** (`qrResult`)

#### スタイル
```css
.qr-section-inline {
    background: #e0f2fe;          /* 水色背景 */
    color: #1e293b;               /* ダークグレーテキスト */
    border-radius: 12px;
    padding: 20px 16px;
    min-width: 240px;
    box-shadow: 0 4px 16px rgba(0,0,0,0.1);
}
```

---

## 🔘 操作ボタン

### 3. QRスキャン操作ボタン（3種類）

#### 3-1. 開始ボタン
```html
<button class="qr-button start" onclick="PickingWork.startQRScan()">開始</button>
```
- **機能**: カメラを起動してQRスキャンを開始
- **スタイル**: 緑色グラデーション (`#10b981` → `#059669`)
- **JavaScript**: `PickingWork.startQRScan()`

#### 3-2. 停止ボタン
```html
<button class="qr-button" onclick="PickingWork.stopQRScan()">停止</button>
```
- **機能**: カメラを停止してスキャンを終了
- **スタイル**: 青色グラデーション (`#3b82f6` → `#1d4ed8`)
- **JavaScript**: `PickingWork.stopQRScan()`

#### 3-3. 手入力ボタン
```html
<button class="qr-button" onclick="PickingWork.manualQRInput()">手入力</button>
```
- **機能**: QRコード値を手動入力（カメラ不要）
- **スタイル**: 青色グラデーション
- **JavaScript**: `PickingWork.manualQRInput()`

---

## 📹 ビデオプレビュー

### 4. カメラプレビュー機能

#### ビデオ要素
```html
<video id="qr-video" playsinline></video>
```

#### スキャンオーバーレイ
```css
.scan-overlay {
    position: absolute;
    top: 50%;
    left: 50%;
    transform: translate(-50%, -50%);
    width: 200px;
    height: 200px;
    border: 2px solid #10b981;  /* 緑色枠 */
    box-shadow: 0 0 0 9999px rgba(0, 0, 0, 0.5);  /* 外側を暗く */
}
```
- **機能**: スキャン範囲を視覚的に表示
- **デザイン**: 200x200pxの緑色枠

#### コンテナ制御
- 初期状態: `display: none` (非表示)
- スキャン開始時: `display: block` (表示)
- スキャン停止時: `display: none` (非表示)

---

## 💻 JavaScript機能

### 5. PickingWork オブジェクト

#### 5-1. プロパティ

```javascript
qrScanner: null,        // QRスキャナーインスタンス
lastQRValue: '',        // 最後にスキャンしたQR値
```

#### 5-2. メソッド一覧

| メソッド名 | 説明 | 引数 | 戻り値 |
|-----------|------|------|--------|
| `startQRScan()` | QRスキャンを開始 | なし | なし |
| `stopQRScan()` | QRスキャンを停止 | なし | なし |
| `handleQRResult(data)` | QR読み取り結果を処理 | data: string | なし |
| `manualQRInput()` | 手動入力ダイアログを表示 | なし | なし |
| `showQRStatus(message, type)` | ステータスメッセージ表示 | message: string, type: string | なし |
| `matchItems()` | QR値と品目IDを照合 | なし | なし |

---

## 🔄 処理フロー

### 6. QRスキャンフロー

#### 6-1. 開始処理 (`startQRScan()`)

```javascript
startQRScan: function() {
    var videoElement = document.getElementById("qr-video");
    var container = document.getElementById('videoContainer');
    
    // 1. ビデオコンテナを表示
    container.style.display = 'block';
    
    // 2. QRスキャナーインスタンス作成
    this.qrScanner = new QrScanner(
        videoElement,
        function(result) {
            PickingWork.handleQRResult(result.data || result);
        },
        {
            returnDetailedScanResult: true,
            highlightScanRegion: false,
            highlightCodeOutline: false
        }
    );
    
    // 3. スキャン開始
    this.qrScanner.start().then(function() {
        PickingWork.showQRStatus('QRコードをスキャン中...', 'info');
    }).catch(function(error) {
        PickingWork.showQRStatus('カメラアクセスエラー: ' + error.message, 'error');
    });
}
```

**処理ステップ:**
1. ビデオコンテナを表示
2. QrScannerインスタンスを作成（コールバック設定）
3. カメラ起動（Promiseベース）
4. 成功時: ステータス表示
5. エラー時: エラーメッセージ表示

#### 6-2. スキャン結果処理 (`handleQRResult(data)`)

```javascript
handleQRResult: function(data) {
    // 1. QR値を保存
    this.lastQRValue = data;
    
    // 2. 結果を画面に表示
    document.getElementById("qrResult").textContent = 'スキャン結果: ' + data;
    
    // 3. スキャンを自動停止
    this.stopQRScan();
    
    // 4. 成功メッセージ表示
    this.showQRStatus('QRコードを読み取りました', 'success');
    
    // 5. 照合処理を自動実行
    this.matchItems();
}
```

**処理ステップ:**
1. QR値を `lastQRValue` に保存
2. 結果表示エリアに値を表示
3. カメラを自動停止
4. 成功ステータス表示（3秒間）
5. 品目照合処理を自動実行

#### 6-3. 停止処理 (`stopQRScan()`)

```javascript
stopQRScan: function() {
    // 1. スキャナーインスタンスが存在する場合
    if (this.qrScanner) {
        this.qrScanner.stop();      // カメラ停止
        this.qrScanner.destroy();   // リソース解放
        this.qrScanner = null;      // 参照クリア
    }
    
    // 2. ビデオコンテナを非表示
    document.getElementById('videoContainer').style.display = 'none';
    
    // 3. ステータス表示
    this.showQRStatus('QRスキャンを停止しました', 'info');
}
```

**処理ステップ:**
1. QrScannerの停止・破棄
2. インスタンス参照をクリア
3. ビデオコンテナを非表示
4. 停止メッセージ表示

---

## 🔍 照合機能

### 7. 品目照合処理 (`matchItems()`)

```javascript
matchItems: function() {
    // 1. 全品目をループ
    PickingWork.currentData.Items.forEach(function(item) {
        // 2. QR値と品目IDを比較
        if (item.Item === PickingWork.lastQRValue) {
            item.sMatching = '一致';
            item.matched = true;
        } else {
            item.sMatching = '不一致';
            item.matched = false;
        }
    });
    
    // 3. 照合結果でソート（一致が上に）
    PickingWork.currentData.Items.sort(function(a, b) {
        if (a.sMatching === b.sMatching) return 0;
        if (a.sMatching === '一致') return -1;
        if (b.sMatching === '一致') return 1;
        return a.sMatching.localeCompare(b.sMatching);
    });
    
    // 4. テーブル再表示
    PickingWork.displayItems(PickingWork.currentData.Items);
}
```

**照合ロジック:**
1. **全品目をループ処理**
2. **QR値と品目ID（Item）を文字列比較**
3. **照合結果を設定**:
   - 一致: `sMatching = '一致'`, `matched = true`
   - 不一致: `sMatching = '不一致'`, `matched = false`
4. **ソート**: 一致項目を上に表示
5. **テーブル再描画**: 照合結果を反映

#### テーブル表示への反映

```javascript
// displayItems() 内での処理
if (detail.matched) {
    row.classList.add('matched');  // 背景色変更（黄色）
}

// 照合結果列の表示
<td><span class="matching-cell">${detail.sMatching || '-'}</span></td>
```

**視覚的フィードバック:**
- 一致行: 背景色が黄色（`background: #fef3c7`）
- 照合結果列: '一致' または '不一致' を表示

---

## 📝 手動入力機能

### 8. 手動QR入力 (`manualQRInput()`)

```javascript
manualQRInput: function() {
    // 1. プロンプトダイアログを表示
    var input = prompt('QRコードの内容を手入力してください:');
    
    // 2. 入力値がある場合
    if (input && input.trim()) {
        // 3. スキャン結果処理を実行（トリム後）
        this.handleQRResult(input.trim());
    }
}
```

**使用ケース:**
- カメラが使えない環境
- QRコードが読み取れない場合
- テスト・デバッグ時

**処理:**
1. ブラウザの `prompt()` ダイアログを表示
2. ユーザー入力を取得
3. トリム（前後の空白除去）
4. `handleQRResult()` に渡して通常のスキャン結果として処理

---

## 💬 ステータスメッセージ

### 9. ステータス表示 (`showQRStatus()`)

```javascript
showQRStatus: function(message, type) {
    var statusDiv = document.getElementById("qrStatusMessage");
    
    // 1. メッセージタイプに応じたクラス設定
    statusDiv.className = "status-message " + type;
    
    // 2. メッセージテキスト設定
    statusDiv.textContent = message;
    
    // 3. 表示
    statusDiv.style.display = 'block';
    
    // 4. 3秒後に自動非表示
    setTimeout(function() {
        statusDiv.style.display = 'none';
    }, 3000);
}
```

#### メッセージタイプ

| タイプ | 色 | 使用場面 |
|--------|-----|----------|
| `info` | 青色 (#3b82f6) | スキャン中、停止時 |
| `success` | 緑色 (#10b981) | 読み取り成功 |
| `error` | 赤色 (#ef4444) | カメラエラー、ライブラリ未読み込み |

#### 表示メッセージ一覧

```javascript
// 開始時
'QRコードをスキャン中...'  // type: info

// 成功時
'QRコードを読み取りました'  // type: success

// 停止時
'QRスキャンを停止しました'  // type: info

// エラー時
'カメラアクセスエラー: [エラーメッセージ]'  // type: error
'QRスキャナーライブラリが読み込まれていません'  // type: error
```

---

## 🔗 C#バックエンド連携

### 10. ItemPicking.cs 関連メソッド

#### 10-1. `addScanTarget()` メソッド

```csharp
/// <summary>
/// QR読み取り対象をGridに追加
/// </summary>
public void addScanTarget(){
    callAPIwithTarget(0, ThisForm.Variables("HiddenValue").Value);
}
```

**機能:**
- QRで読み取った値を使用してAPIを呼び出し
- Gridに検索結果を追加表示

**使用変数:**
- `HiddenValue`: QRスキャン値を格納

#### 10-2. `matchGridItems()` メソッド

```csharp
/// <summary>
/// Grid内にある品目IDを照合
/// </summary>
public void matchGridItems(){
    int count = ThisForm.Components["ResultGrid"].GetGridRowCount();
    if(count == 0) return;
    
    // HiddenValueが空の場合はgItemを使用
    if(ThisForm.Variables("HiddenValue").Value == "")
        ThisForm.Variables("HiddenValue").Value = ThisForm.Variables("gItem").Value;
    
    // 全行をループして照合
    for(int i = 1; i < count+1; i++){
        if(ThisForm.Components["ResultGrid"].GetGridValue(i,6) == ThisForm.Variables("HiddenValue").Value){
            ThisForm.Components["ResultGrid"].SetGridValue(i,8,"OK");  // 一致
        }
        else{
            ThisForm.Components["ResultGrid"].SetGridValue(i,8,"NG");  // 不一致
        }
    }
}
```

**機能:**
- ResultGrid の全行をスキャン
- 6列目（品目ID）とQR値を比較
- 8列目（照合結果）に "OK" または "NG" を設定

**照合ロジック:**
1. Grid行数取得
2. HiddenValue（QR値）取得
3. 全行をループ
4. 品目ID（6列目）と比較
5. 照合結果（8列目）に結果設定

---

## 📊 データフロー

### 11. QRスキャンからデータ保存までのフロー

```
[ユーザー操作]
    ↓
[開始ボタンクリック]
    ↓
[startQRScan()]
    ↓
[カメラ起動] ← QrScanner.start()
    ↓
[QRコード読み取り]
    ↓
[handleQRResult(data)]
    ↓
    ├─ lastQRValue = data（保存）
    ├─ stopQRScan()（自動停止）
    └─ matchItems()（照合実行）
        ↓
        ├─ Items全体をループ
        ├─ Item === lastQRValue で比較
        ├─ sMatching, matched を設定
        ├─ 照合結果でソート
        └─ displayItems()（再描画）
            ↓
[テーブル表示更新]
    ├─ 一致行は黄色背景
    └─ 照合結果列に '一致'/'不一致' 表示
    ↓
[ユーザーがピック数入力]
    ↓
[保存ボタンクリック]
    ↓
[savePickingData()]
    ↓
[C#バックエンドへ送信]
    └─ WSForm.generate('ReturnToDetail')
```

---

## 🎨 スタイル定義

### 12. CSS クラス一覧

#### QRセクション関連

```css
/* QRスキャンセクション */
.qr-section-inline {
    background: #e0f2fe;
    color: #1e293b;
    border-radius: 12px;
    padding: 20px 16px;
    min-width: 240px;
    box-shadow: 0 4px 16px rgba(0,0,0,0.1);
}

/* QRタイトル */
.qr-title {
    font-size: 18px;
    font-weight: 600;
    margin-bottom: 16px;
}

/* QRコントロールボタングループ */
.qr-controls {
    display: flex;
    gap: 8px;
    margin-bottom: 16px;
    flex-wrap: wrap;
}

/* QRボタン */
.qr-button {
    background: linear-gradient(135deg, #3b82f6 0%, #1d4ed8 100%);
    color: white;
    border: none;
    padding: 10px 16px;
    border-radius: 6px;
    cursor: pointer;
    transition: all 0.2s ease;
    box-shadow: 0 2px 8px rgba(59, 130, 246, 0.3);
}

.qr-button:hover {
    transform: translateY(-2px);
    box-shadow: 0 4px 12px rgba(59, 130, 246, 0.4);
}

/* 開始ボタン（緑色） */
.qr-button.start {
    background: linear-gradient(135deg, #10b981 0%, #059669 100%);
    box-shadow: 0 2px 8px rgba(16, 185, 129, 0.3);
}

.qr-button.start:hover {
    box-shadow: 0 4px 12px rgba(16, 185, 129, 0.4);
}

/* ビデオコンテナ */
#qr-video {
    width: 100%;
    max-width: 320px;
    border-radius: 8px;
    display: block;
    margin: 0 auto;
}

/* スキャンオーバーレイ */
.scan-overlay {
    position: absolute;
    top: 50%;
    left: 50%;
    transform: translate(-50%, -50%);
    width: 200px;
    height: 200px;
    border: 2px solid #10b981;
    box-shadow: 0 0 0 9999px rgba(0, 0, 0, 0.5);
}

/* QR結果表示 */
.qr-result {
    background: white;
    padding: 12px;
    border-radius: 6px;
    text-align: center;
    font-weight: 500;
    margin-top: 12px;
    color: #1e293b;
}
```

#### ステータスメッセージ

```css
.status-message {
    padding: 12px;
    border-radius: 6px;
    margin-bottom: 12px;
    font-weight: 500;
}

.status-message.info {
    background: #dbeafe;
    color: #1e40af;
    border: 1px solid #3b82f6;
}

.status-message.success {
    background: #d1fae5;
    color: #065f46;
    border: 1px solid #10b981;
}

.status-message.error {
    background: #fee2e2;
    color: #991b1b;
    border: 1px solid #ef4444;
}
```

#### 照合結果行

```css
/* 一致行のハイライト */
.matched {
    background: #fef3c7 !important;
    font-weight: 600;
}

.matched .matching-cell {
    color: #059669;
    font-weight: 700;
}
```

---

## 🔧 設定オプション

### 13. QrScanner 初期化オプション

```javascript
new QrScanner(
    videoElement,
    callbackFunction,
    {
        returnDetailedScanResult: true,    // 詳細結果を返す
        highlightScanRegion: false,        // スキャン領域ハイライト無効
        highlightCodeOutline: false        // QRコード輪郭ハイライト無効
    }
);
```

| オプション | 値 | 説明 |
|-----------|-----|------|
| `returnDetailedScanResult` | `true` | result.data でQR値を取得 |
| `highlightScanRegion` | `false` | ライブラリのハイライト機能を無効化 |
| `highlightCodeOutline` | `false` | QRコード輪郭描画を無効化 |

**理由:** カスタムの `scan-overlay` を使用するため、ライブラリのハイライト機能は不要

---

## 🌐 外部公開API

### 14. グローバル関数（CSI連携用）

```javascript
// CSIから呼び出し可能な関数
window.startQRScanning = function() {
    PickingWork.startQRScan();
};

window.executeMatching = function() {
    PickingWork.performMatching();
};

window.setPickingDetail = function(data) {
    PickingWork.currentData = data;
    PickingWork.displayData(data);
};

window.clearPickingSelection = function() {
    PickingWork.clearSelection();
};
```

**用途:** C#バックエンドから JavaScript 関数を直接呼び出し

---

## 📱 レスポンシブ対応

### 15. メディアクエリ

```css
/* 画面幅260px以下 */
@media (max-width: 260px) {
    .header-qr-row {
        flex-direction: column;  /* 縦並びに変更 */
        gap: 12px;
    }
    
    .qr-section-inline, .item-header {
        max-width: 100%;
        min-width: 0;
        width: 100%;
    }
}
```

**動作:**
- 通常: 品目情報とQRセクションが横並び
- 狭い画面: 縦並びに自動切り替え

---

## 🔐 セキュリティ考慮事項

### 16. カメラアクセス許可

- **HTTPS必須**: QR-Scannerはセキュアコンテキストでのみ動作
- **ユーザー許可**: ブラウザがカメラアクセス許可を要求
- **エラーハンドリング**: 許可拒否時のエラーメッセージ表示

```javascript
this.qrScanner.start().catch(function(error) {
    PickingWork.showQRStatus('カメラアクセスエラー: ' + error.message, 'error');
});
```

---

## 📈 パフォーマンス

### 17. 最適化ポイント

1. **自動停止**: スキャン成功後は自動的にカメラ停止
2. **リソース解放**: `destroy()` でメモリリーク防止
3. **CDN利用**: QR-Scannerライブラリをキャッシュ可能
4. **遅延初期化**: DOMContentLoaded後に初期化

```javascript
document.addEventListener('DOMContentLoaded', function() {
    setTimeout(() => {
        PickingWork.init();
    }, 200);  // 200ms遅延で安全に初期化
});
```

---

## 🐛 トラブルシューティング

### 18. よくある問題と解決策

| 問題 | 原因 | 解決策 |
|------|------|--------|
| カメラが起動しない | HTTPS未使用 | HTTPS環境で実行 |
| ライブラリ未読み込みエラー | CDN接続失敗 | インターネット接続確認 |
| QRコードが読み取れない | 照明不足、距離 | カメラに近づける、照明改善 |
| スキャン後も動作し続ける | 自動停止失敗 | 手動で停止ボタン押下 |
| 照合結果が表示されない | データ不整合 | ブラウザコンソールでエラー確認 |

---

## 🎯 今後の拡張可能性

### 19. 機能拡張案

1. **複数QRコード対応**
   - 連続スキャンモード
   - スキャン履歴保存

2. **バーコード対応**
   - 1D/2Dバーコード読み取り
   - JAN/EANコード対応

3. **カメラ選択**
   - フロント/リアカメラ切り替え
   - デバイスリスト表示

4. **スキャン精度向上**
   - ズーム機能
   - フラッシュライト制御

5. **音声フィードバック**
   - スキャン成功時のビープ音
   - エラー時の警告音

---

## 📚 参考情報

### 20. 関連リソース

- **QR-Scanner GitHub**: https://github.com/nimiq/qr-scanner
- **npm パッケージ**: https://www.npmjs.com/package/qr-scanner
- **CDN**: https://unpkg.com/qr-scanner@1.4.2/

### ブラウザ対応

| ブラウザ | 対応状況 | 備考 |
|---------|---------|------|
| Chrome | ✅ | 完全対応 |
| Firefox | ✅ | 完全対応 |
| Safari | ✅ | iOS 11+ |
| Edge | ✅ | Chromium版 |
| IE | ❌ | 非対応 |

---

## 📝 まとめ

### 主要機能
✅ QR-Scanner 1.4.2 ライブラリ使用  
✅ カメラベースのリアルタイムスキャン  
✅ 手動入力フォールバック  
✅ 自動照合・ソート機能  
✅ 視覚的フィードバック（ステータスメッセージ、ハイライト）  
✅ レスポンシブデザイン  
✅ C#バックエンド連携  

### コード配置
- **HTML**: ItemPicking.html (12行目、420-1009行目)
- **JavaScript**: PickingWork オブジェクト内（830-950行目）
- **C#**: ItemPicking.cs (338-356行目)

---

**作成日**: 2025年10月14日  
**バージョン**: 1.0  
**QR-Scanner**: v1.4.2
