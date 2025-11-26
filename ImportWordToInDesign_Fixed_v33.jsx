// ImportWordToInDesign_Fixed_v33.jsx
// 小項目に□、リストに・を段落先頭に追加(箇条書き機能不使用)
// 文字サイズを14.817Qに統一
// 番号リストの自動採番機能を追加
// 演習タイトルの後で番号リストをリセット

#target indesign

// グローバル設定
var CONFIG = {
    autoCreatePages: true,
    maxSpreads: 100,
    
    // 使用するマスターページ
    masterPagePrefix: "H",
    masterPageName: "本文マスター",
    
    // フォント設定
    defaultFont: {
        family: "BIZ UDゴシック",
        style: "Regular"
    },
    
    // 文字サイズ設定(Q単位)
    fontSize: 14.817,
    
    // マージン設定(フォールバック用)
    fallbackMargin: {
        top: 36,
        bottom: 36,
        inside: 54,
        outside: 36
    },
    
    // スタイルマッピング
    styleMapping: {
        "Heading 1": "大見出し1",
        "Heading 2": "大見出し1",
        "Heading 3": "見出し",
        "大項目": "大見出し1",
        "小項目": "小項目",
        "リスト": "リスト",
        "箇条書き": "リスト",
        "List Paragraph": "リスト",
        "リスト段落": "リスト段落",
        "Normal": "Normal",
        "BodyText": "Normal",
        "演習タイトル": "演習タイトル"
    },
    
    // 段落先頭記号
    bulletSymbols: {
        kokomoku: "□　", // 小項目用(□+全角スペース)
        list: "・　"      // リスト用(・+全角スペース)
    },
    
    // 番号リスト設定
    numberedList: {
        // Wordの番号リストスタイル名(実際の名前に合わせて調整)
        styleNames: ["演習", "手順", "番号付きリスト", "List Number"],
        useAutoNumbering: true,  // InDesignの自動採番を使用
        format: "^#.^t",         // 形式: 1. 2. 3. (^#=番号、^t=タブ)
        startAt: 1,
        // 演習タイトルの後でリセット
        resetAfterStyle: "演習タイトル"
    },
    
    // ハイパーリンクの色設定
    hyperlinkColor: {
        applyBlack: true,
        removeUnderline: false
    },

    // 表（テーブル）設定
    tableSettings: {
        enabled: true,
        // 罫線設定
        strokeWeight: 0.5,
        strokeColor: "Black",
        // セル設定
        cellInset: {
            top: 2,
            bottom: 2,
            left: 3,
            right: 3
        },
        // ヘッダー行の設定
        headerRowFill: "None",
        headerRowBold: true,
        // 表のフォント設定（CONFIG.defaultFontを使用）
        applyDefaultFont: true
    },
    
    // スタイルを強制的に再作成
    forceRecreateStyles: false,
    
    // デバッグモード
    debugMode: true
};

// デバッグログ
function debugLog(message) {
    if (CONFIG.debugMode) {
        $.writeln("[DEBUG] " + message);
    }
}

// メイン処理
function main() {
    if (app.documents.length === 0) {
        alert("InDesignドキュメントを開いてからスクリプトを実行してください。");
        return;
    }
    
    var doc = app.activeDocument;
    
    // H-本文マスターの存在確認
    var hMaster = getMasterPage(doc, CONFIG.masterPagePrefix + "-" + CONFIG.masterPageName);
    if (!hMaster) {
        alert("エラー: 「" + CONFIG.masterPagePrefix + "-" + CONFIG.masterPageName + "」マスターページが見つかりません。");
        return;
    }
    
    debugLog("使用マスター: " + hMaster.name);
    
    // Word文書を選択
    var wordFile = File.openDialog("インポートするWord文書を選択してください (.docx)", "*.docx");
    if (!wordFile) {
        return;
    }
    
    // 確認ダイアログ
    var confirmMsg = "Word文書をインポートします\n\n";
    confirmMsg += "【段落スタイル変換】\n";
    confirmMsg += "・大項目 → 大見出し1\n";
    confirmMsg += "・小項目 → 小項目(太字、14.817Q、□付き)\n";
    confirmMsg += "・リスト → リスト(14.817Q、・付き)\n";
    confirmMsg += "・番号リスト → 自動採番(演習タイトル後でリセット)\n\n";
    confirmMsg += "【表（テーブル）処理】\n";
    confirmMsg += "・Word内の表を保持・フォーマット\n";
    confirmMsg += "・罫線・セル余白の統一\n\n";
    confirmMsg += "【フォント統一】\n";
    confirmMsg += "・BIZ UDゴシックに統一\n";
    confirmMsg += "・MS明朝などを強制置換\n\n";
    confirmMsg += "実行しますか?";
    
    if (!confirm(confirmMsg)) {
        return;
    }
    
    try {
        app.scriptPreferences.enableRedraw = false;
        
        var startTime = new Date();
        
        // 必要なスタイルを作成・更新
        createRequiredStyles(doc);
        
        // 現在のページ数を記録
        var initialPageCount = doc.pages.length;
        
        // Word文書をインポート
        var result = importWordDocument(doc, wordFile, hMaster);
        
        // MS明朝などの不要フォントを強制置換
        result.fontsReplaced = replaceAllFonts(doc);
        
        // 番号リストを修正(演習タイトル後でリセット)
        result.numberedListFixed = fixNumberedListsWithReset(doc);
        
        // 小項目とリストの先頭に記号を追加
        result.kokomokuFixed = addBulletToKokomoku(doc);
        result.listFixed = addBulletToList(doc);
        
        // 小項目とリストの書式を文字レベルで強制適用
        result.kokomokuFormatted = fixKokomokuFormatting(doc);
        result.listFormatted = fixListFormatting(doc);
        
        // ハイパーリンクの色を変更
        result.hyperlinksFixed = fixHyperlinkColors(doc);

        // Wordからインポートした表を処理
        var tableResult = processImportedTables(doc);
        result.tablesProcessed = tableResult.tablesProcessed;
        result.cellsProcessed = tableResult.cellsProcessed;

        var endTime = new Date();
        var duration = (endTime - startTime) / 1000;
        
        app.scriptPreferences.enableRedraw = true;
        
        var resultMsg = "完了!\n\n";
        resultMsg += "既存ページ: " + initialPageCount + "p\n";
        resultMsg += "追加見開き: " + result.spreadsCreated + "\n";
        resultMsg += "段落数: " + result.paragraphsImported + "\n";
        resultMsg += "スタイル適用: " + result.stylesApplied + "件\n";
        resultMsg += "フォント置換: " + result.fontsReplaced + "件\n";
        resultMsg += "番号リスト修正: " + result.numberedListFixed + "件\n";
        resultMsg += "小項目記号追加: " + result.kokomokuFixed + "件\n";
        resultMsg += "リスト記号追加: " + result.listFixed + "件\n";
        resultMsg += "リンク修正: " + result.hyperlinksFixed + "件\n";
        resultMsg += "表処理: " + result.tablesProcessed + "表 (" + result.cellsProcessed + "セル)\n";
        resultMsg += "処理時間: " + duration.toFixed(1) + "秒";
        
        alert(resultMsg);
        
    } catch (e) {
        app.scriptPreferences.enableRedraw = true;
        alert("エラー:\n\n" + e.message + "\n\n行: " + e.line);
    }
}

// マスターページを取得
function getMasterPage(doc, masterName) {
    try {
        var master = doc.masterSpreads.itemByName(masterName);
        if (master.isValid) {
            return master;
        }
    } catch (e) {}
    return null;
}

// 必要なスタイルを作成・更新
function createRequiredStyles(doc) {
    debugLog("=== スタイル作成・更新開始 ===");
    
    updateKokomokuStyle(doc);
    updateListStyle(doc);
    createNumberedListStyle(doc);
    createEnshuTitleStyle(doc);
    
    debugLog("=== スタイル作成・更新完了 ===");
}

// 「小項目」スタイルを作成・更新(箇条書き機能不使用)
function updateKokomokuStyle(doc) {
    var styleName = "小項目";
    var style = doc.paragraphStyles.itemByName(styleName);
    
    if (style.isValid && CONFIG.forceRecreateStyles) {
        debugLog("既存の「小項目」スタイルを削除します");
        try {
            style.remove(doc.paragraphStyles.item(0));
        } catch (e) {}
    }
    
    style = doc.paragraphStyles.itemByName(styleName);
    if (!style.isValid) {
        style = doc.paragraphStyles.add({name: styleName});
    }
    
    // 基本設定
    style.appliedFont = CONFIG.defaultFont.family;
    style.fontStyle = "Bold";
    style.pointSize = CONFIG.fontSize;
    style.spaceBefore = 6;
    style.spaceAfter = 3;
    
    // インデント設定(ぶら下げインデント)
    style.leftIndent = 14.17;
    style.firstLineIndent = -14.17;
    
    // 箇条書き機能は使用しない
    style.bulletsAndNumberingListType = ListType.NO_LIST;
    
    debugLog("「小項目」スタイル設定完了(14.817Q)");
}

// 「リスト」スタイルを作成・更新(箇条書き機能不使用)
function updateListStyle(doc) {
    var styleName = "リスト";
    var style = doc.paragraphStyles.itemByName(styleName);
    
    if (style.isValid && CONFIG.forceRecreateStyles) {
        debugLog("既存の「リスト」スタイルを削除します");
        try {
            style.remove(doc.paragraphStyles.item(0));
        } catch (e) {}
    }
    
    style = doc.paragraphStyles.itemByName(styleName);
    if (!style.isValid) {
        style = doc.paragraphStyles.add({name: styleName});
    }
    
    // 基本設定
    style.appliedFont = CONFIG.defaultFont.family;
    style.fontStyle = "Regular";
    style.pointSize = CONFIG.fontSize;
    style.spaceBefore = 3;
    style.spaceAfter = 3;
    
    // インデント設定(ぶら下げインデント)
    style.leftIndent = 10;
    style.firstLineIndent = -10;
    
    // 箇条書き機能は使用しない
    style.bulletsAndNumberingListType = ListType.NO_LIST;
    
    debugLog("「リスト」スタイル設定完了(14.817Q)");
}

// 「演習タイトル」スタイルを作成
function createEnshuTitleStyle(doc) {
    var styleName = "演習タイトル";
    var style = doc.paragraphStyles.itemByName(styleName);
    
    if (!style.isValid) {
        style = doc.paragraphStyles.add({name: styleName});
        
        // 基本設定(必要に応じて調整)
        style.appliedFont = CONFIG.defaultFont.family;
        style.fontStyle = "Bold";
        style.pointSize = CONFIG.fontSize * 1.2; // 少し大きめ
        style.spaceBefore = 12;
        style.spaceAfter = 6;
        
        debugLog("「演習タイトル」スタイル作成完了");
    }
}

// 番号リストスタイルを作成
function createNumberedListStyle(doc) {
    var styleName = "番号リスト";
    var style = doc.paragraphStyles.itemByName(styleName);
    
    if (style.isValid && CONFIG.forceRecreateStyles) {
        debugLog("既存の「番号リスト」スタイルを削除します");
        try {
            style.remove(doc.paragraphStyles.item(0));
        } catch (e) {}
    }
    
    style = doc.paragraphStyles.itemByName(styleName);
    if (!style.isValid) {
        style = doc.paragraphStyles.add({name: styleName});
    }
    
    // 基本設定
    style.appliedFont = CONFIG.defaultFont.family;
    style.fontStyle = "Regular";
    style.pointSize = CONFIG.fontSize;
    style.spaceBefore = 3;
    style.spaceAfter = 3;
    
    // インデント設定
    style.leftIndent = 14.17;
    style.firstLineIndent = -14.17;
    
    if (CONFIG.numberedList.useAutoNumbering) {
        // 自動採番を使用
        style.bulletsAndNumberingListType = ListType.NUMBERED_LIST;
        // numberingFormat削除（互換性問題）
        style.numberingStartAt = CONFIG.numberedList.startAt;
        style.numberingContinue = false; // 明示的にfalseに設定
        
        debugLog("番号リストスタイル設定完了(自動採番、リセット機能付き)");
    } else {
        style.bulletsAndNumberingListType = ListType.NO_LIST;
        debugLog("番号リストスタイル設定完了(手動)");
    }
}

// 番号リストを検出して修正(演習タイトル後でリセット)
function fixNumberedListsWithReset(doc) {
    var fixCount = 0;
    var resetCount = 0;
    
    debugLog("=== 番号リストの検出・修正開始(リセット機能付き) ===");
    
    var numberedListStyle = doc.paragraphStyles.itemByName("番号リスト");
    if (!numberedListStyle.isValid) {
        debugLog("番号リストスタイルが見つかりません");
        return 0;
    }
    
    var enshuTitleStyle = doc.paragraphStyles.itemByName("演習タイトル");
    var stories = doc.stories;
    
    for (var i = 0; i < stories.length; i++) {
        var paragraphs = stories[i].paragraphs;
        var shouldRestart = false;
        
        for (var j = 0; j < paragraphs.length; j++) {
            var para = paragraphs[j];
            var paraText = para.contents;
            var currentStyleName = para.appliedParagraphStyle.name;
            
            // 演習タイトルを検出
            if (enshuTitleStyle.isValid && currentStyleName === "演習タイトル") {
                shouldRestart = true;
                debugLog("演習タイトル検出: 段落 " + j + " - 次の番号リストをリセット");
                continue;
            }
            
            // 行頭が数字で始まる段落を検出
            var numberPattern = /^(\d+)[.)\s　]/;
            var match = paraText.match(numberPattern);
            
            if (match) {
                try {
                    // 番号リストスタイルを適用
                    para.appliedParagraphStyle = numberedListStyle;
                    
                    // 演習タイトルの直後の場合は番号を再開
                    if (shouldRestart) {
                        para.numberingContinue = false; // 番号をリセット
                        para.numberingStartAt = 1;
                        shouldRestart = false;
                        resetCount++;
                        debugLog("番号リストをリセット: 段落 " + j);
                    } else {
                        para.numberingContinue = true; // 番号を継続
                    }
                    
                    // 既存の番号を削除(自動採番を使う場合)
                    if (CONFIG.numberedList.useAutoNumbering) {
                        var numLength = match[0].length;
                        para.contents = paraText.substring(numLength);
                    }
                    
                    fixCount++;
                    
                    if (fixCount <= 5) {
                        debugLog("番号リスト検出: 段落 " + j + " - " + match[0]);
                    }
                } catch (e) {
                    debugLog("番号リスト適用エラー: " + e.message);
                }
            }
        }
    }
    
    debugLog("番号リスト修正完了: " + fixCount + "件 (リセット: " + resetCount + "回)");
    return fixCount;
}

// Word文書をインポート
function importWordDocument(doc, wordFile, master) {
    var result = {
        spreadsCreated: 0,
        paragraphsImported: 0,
        stylesApplied: 0,
        fontsReplaced: 0,
        numberedListFixed: 0,
        kokomokuFixed: 0,
        listFixed: 0,
        kokomokuFormatted: 0,
        listFormatted: 0,
        hyperlinksFixed: 0,
        tablesProcessed: 0,
        cellsProcessed: 0
    };
    
    var lastPage = doc.pages[doc.pages.length - 1];
    var textFrame = getOrCreateTextFrame(doc, lastPage);
    
    if (!textFrame) {
        throw new Error("テキストフレームを作成できませんでした。");
    }
    
    try {
        textFrame.place(wordFile);
    } catch (e) {
        throw new Error("Word文書の読み込み失敗: " + e.message);
    }
    
    if (textFrame.overflows && CONFIG.autoCreatePages) {
        result.spreadsCreated = autoFlowWithSpreads(doc, textFrame, master);
    }
    
    result.stylesApplied = applyStyleMapping(doc);
    result.paragraphsImported = countParagraphs(doc);
    
    return result;
}

// 全フォントをBIZ UDゴシックに強制置換
function replaceAllFonts(doc) {
    var replaceCount = 0;
    var targetFont, targetBoldFont;
    
    try {
        targetFont = app.fonts.itemByName(CONFIG.defaultFont.family + "\t" + CONFIG.defaultFont.style);
    } catch (e) {
        try {
            targetFont = app.fonts.itemByName(CONFIG.defaultFont.family);
        } catch (e2) {
            return 0;
        }
    }
    
    try {
        targetBoldFont = app.fonts.itemByName(CONFIG.defaultFont.family + "\tBold");
    } catch (e) {
        targetBoldFont = targetFont;
    }
    
    var stories = doc.stories;
    
    for (var i = 0; i < stories.length; i++) {
        try {
            var texts = stories[i].texts;
            
            for (var j = 0; j < texts.length; j++) {
                try {
                    var currentFontName = texts[j].appliedFont.name;
                    
                    if (currentFontName.indexOf("MS") === 0 || currentFontName.indexOf("ＭＳ") === 0) {
                        var isBold = (currentFontName.indexOf("Bold") >= 0 || 
                                     currentFontName.indexOf("太字") >= 0);
                        
                        texts[j].appliedFont = isBold ? targetBoldFont : targetFont;
                        replaceCount++;
                    }
                } catch (e) {}
            }
        } catch (e) {}
    }
    
    return replaceCount;
}

// 小項目の段落先頭に□記号を追加
function addBulletToKokomoku(doc) {
    var addCount = 0;
    var kokomokuStyle = doc.paragraphStyles.itemByName("小項目");
    
    if (!kokomokuStyle.isValid) {
        return 0;
    }
    
    debugLog("=== 小項目に□記号を追加開始 ===");
    
    var stories = doc.stories;
    var bulletSymbol = CONFIG.bulletSymbols.kokomoku;
    
    for (var i = 0; i < stories.length; i++) {
        var paragraphs = stories[i].paragraphs;
        
        for (var j = 0; j < paragraphs.length; j++) {
            var para = paragraphs[j];
            
            if (para.appliedParagraphStyle.name === "小項目") {
                try {
                    var paraText = para.contents;
                    
                    // 既に記号がある場合はスキップ
                    if (paraText.indexOf(bulletSymbol) === 0) {
                        continue;
                    }
                    
                    // 段落の先頭に記号を挿入
                    para.contents = bulletSymbol + paraText;
                    addCount++;
                    
                    if (addCount <= 3) {
                        debugLog("小項目記号追加: 段落 " + j);
                    }
                } catch (e) {
                    debugLog("小項目記号追加エラー: " + e.message);
                }
            }
        }
    }
    
    debugLog("小項目記号追加完了: " + addCount + "件");
    return addCount;
}

// リストの段落先頭に・記号を追加
function addBulletToList(doc) {
    var addCount = 0;
    var listStyle = doc.paragraphStyles.itemByName("リスト");
    
    if (!listStyle.isValid) {
        return 0;
    }
    
    debugLog("=== リストに・記号を追加開始 ===");
    
    var stories = doc.stories;
    var bulletSymbol = CONFIG.bulletSymbols.list;
    
    for (var i = 0; i < stories.length; i++) {
        var paragraphs = stories[i].paragraphs;
        
        for (var j = 0; j < paragraphs.length; j++) {
            var para = paragraphs[j];
            
            if (para.appliedParagraphStyle.name === "リスト") {
                try {
                    var paraText = para.contents;
                    
                    // 既に記号がある場合はスキップ
                    if (paraText.indexOf(bulletSymbol) === 0) {
                        continue;
                    }
                    
                    // 段落の先頭に記号を挿入
                    para.contents = bulletSymbol + paraText;
                    addCount++;
                    
                    if (addCount <= 3) {
                        debugLog("リスト記号追加: 段落 " + j);
                    }
                } catch (e) {
                    debugLog("リスト記号追加エラー: " + e.message);
                }
            }
        }
    }
    
    debugLog("リスト記号追加完了: " + addCount + "件");
    return addCount;
}

// 小項目の書式を文字レベルで強制適用
function fixKokomokuFormatting(doc) {
    var fixCount = 0;
    var kokomokuStyle = doc.paragraphStyles.itemByName("小項目");
    
    if (!kokomokuStyle.isValid) {
        return 0;
    }
    
    debugLog("=== 小項目の書式修正開始 ===");
    
    var boldFont;
    try {
        boldFont = app.fonts.itemByName(CONFIG.defaultFont.family + "\tBold");
    } catch (e) {
        return 0;
    }
    
    var stories = doc.stories;
    
    for (var i = 0; i < stories.length; i++) {
        var paragraphs = stories[i].paragraphs;
        
        for (var j = 0; j < paragraphs.length; j++) {
            var para = paragraphs[j];
            
            if (para.appliedParagraphStyle.name === "小項目") {
                try {
                    // 文字レベルで適用
                    var chars = para.characters;
                    for (var k = 0; k < chars.length; k++) {
                        chars[k].appliedFont = boldFont;
                        chars[k].pointSize = CONFIG.fontSize;
                    }
                    
                    fixCount++;
                    
                    if (fixCount <= 3) {
                        debugLog("小項目書式修正: 段落 " + j + " (文字数: " + chars.length + ")");
                    }
                } catch (e) {
                    debugLog("小項目書式修正エラー: " + e.message);
                }
            }
        }
    }
    
    debugLog("小項目書式修正完了: " + fixCount + "件");
    return fixCount;
}

// リストの書式を文字レベルで強制適用
function fixListFormatting(doc) {
    var fixCount = 0;
    var listStyle = doc.paragraphStyles.itemByName("リスト");
    
    if (!listStyle.isValid) {
        return 0;
    }
    
    debugLog("=== リストの書式修正開始 ===");
    
    var regularFont;
    try {
        regularFont = app.fonts.itemByName(CONFIG.defaultFont.family + "\t" + CONFIG.defaultFont.style);
    } catch (e) {
        try {
            regularFont = app.fonts.itemByName(CONFIG.defaultFont.family);
        } catch (e2) {
            return 0;
        }
    }
    
    var stories = doc.stories;
    
    for (var i = 0; i < stories.length; i++) {
        var paragraphs = stories[i].paragraphs;
        
        for (var j = 0; j < paragraphs.length; j++) {
            var para = paragraphs[j];
            
            if (para.appliedParagraphStyle.name === "リスト") {
                try {
                    // 文字レベルで適用
                    var chars = para.characters;
                    for (var k = 0; k < chars.length; k++) {
                        chars[k].appliedFont = regularFont;
                        chars[k].pointSize = CONFIG.fontSize;
                    }
                    
                    fixCount++;
                    
                    if (fixCount <= 3) {
                        debugLog("リスト書式修正: 段落 " + j + " (文字数: " + chars.length + ")");
                    }
                } catch (e) {
                    debugLog("リスト書式修正エラー: " + e.message);
                }
            }
        }
    }
    
    debugLog("リスト書式修正完了: " + fixCount + "件");
    return fixCount;
}

// テキストフレームを取得または作成
function getOrCreateTextFrame(doc, page) {
    // まずマスターページのテキストフレームをオーバーライド
    var masterFrame = overrideMasterTextFrame(page);
    if (masterFrame) {
        return masterFrame;
    }
    
    // マスターフレームがない場合、ページ上の最大のテキストフレームを探す
    var textFrames = page.textFrames;
    var largestFrame = null;
    var largestArea = 0;
    
    for (var i = 0; i < textFrames.length; i++) {
        // 既に連結されているフレームはスキップ
        if (textFrames[i].previousTextFrame) {
            continue;
        }
        
        var bounds = textFrames[i].geometricBounds;
        var area = (bounds[2] - bounds[0]) * (bounds[3] - bounds[1]);
        
        if (area > largestArea) {
            largestArea = area;
            largestFrame = textFrames[i];
        }
    }
    
    if (largestFrame) {
        return largestFrame;
    }
    
    // テキストフレームがない場合は新規作成
    return createTextFrame(doc, page);
}

// マスターページのテキストフレームをオーバーライド
function overrideMasterTextFrame(page) {
    try {
        var appliedMaster = page.appliedMaster;
        
        if (!appliedMaster || appliedMaster.name === "None") {
            return null;
        }
        
        try {
            page.overrideAll();
            
            var textFrames = page.textFrames;
            
            for (var i = 0; i < textFrames.length; i++) {
                if (!textFrames[i].previousTextFrame) {
                    return textFrames[i];
                }
            }
            
            if (textFrames.length > 0) {
                return textFrames[0];
            }
        } catch (e) {}
        
        var masterPages = appliedMaster.pages;
        
        for (var i = 0; i < masterPages.length; i++) {
            var masterPage = masterPages[i];
            
            if (masterPage.side === page.side || masterPage.side === PageSideOptions.SINGLE_SIDED) {
                var masterTextFrames = masterPage.textFrames;
                
                if (masterTextFrames.length > 0) {
                    try {
                        return masterTextFrames[0].override(page);
                    } catch (e) {}
                }
            }
        }
        
    } catch (e) {}
    
    return null;
}

// テキストフレームを新規作成
function createTextFrame(doc, page) {
    var pageWidth = doc.documentPreferences.pageWidth;
    var pageHeight = doc.documentPreferences.pageHeight;
    var margin = CONFIG.fallbackMargin;
    
    var isLeftPage = (page.side === PageSideOptions.LEFT_HAND);
    var isFacingPages = doc.documentPreferences.facingPages;
    
    var left, right;
    
    if (isFacingPages) {
        if (isLeftPage) {
            left = margin.inside;
            right = pageWidth - margin.outside;
        } else {
            left = margin.outside;
            right = pageWidth - margin.inside;
        }
    } else {
        left = margin.outside;
        right = pageWidth - margin.inside;
    }
    
    var top = margin.top;
    var bottom = pageHeight - margin.bottom;
    var bounds = [top, left, bottom, right];
    
    return page.textFrames.add({geometricBounds: bounds});
}

// 見開き単位でページを自動追加
function autoFlowWithSpreads(doc, startFrame, master) {
    var spreadsCreated = 0;
    var currentFrame = startFrame;
    var loopCount = 0;
    
    while (currentFrame.overflows && loopCount < CONFIG.maxSpreads) {
        loopCount++;
        
        var newSpread = doc.spreads.add(LocationOptions.AT_END);
        spreadsCreated++;
        
        try {
            newSpread.appliedMaster = master;
        } catch (e) {}
        
        var pages = newSpread.pages;
        
        if (pages.length === 0) {
            break;
        }
        
        for (var i = 0; i < pages.length; i++) {
            var pageFrame = getOrCreateTextFrame(doc, pages[i]);
            
            if (pageFrame) {
                try {
                    pageFrame.previousTextFrame = currentFrame;
                    currentFrame = pageFrame;
                    
                    if (!currentFrame.overflows || i === pages.length - 1) {
                        break;
                    }
                } catch (e) {
                    return spreadsCreated;
                }
            }
        }
        
        while (currentFrame.nextTextFrame) {
            currentFrame = currentFrame.nextTextFrame;
        }
    }
    
    return spreadsCreated;
}

// スタイルマッピングを適用
function applyStyleMapping(doc) {
    var stories = doc.stories;
    var mappingCount = 0;
    
    for (var i = 0; i < stories.length; i++) {
        var paragraphs = stories[i].paragraphs;
        
        for (var j = 0; j < paragraphs.length; j++) {
            var para = paragraphs[j];
            var currentStyleName = para.appliedParagraphStyle.name;
            
            if (CONFIG.styleMapping.hasOwnProperty(currentStyleName)) {
                var targetStyleName = CONFIG.styleMapping[currentStyleName];
                
                try {
                    var targetStyle = doc.paragraphStyles.itemByName(targetStyleName);
                    if (targetStyle.isValid) {
                        para.appliedParagraphStyle = targetStyle;
                        mappingCount++;
                    }
                } catch (e) {}
            }
        }
    }
    
    return mappingCount;
}

// ハイパーリンクの色を修正
function fixHyperlinkColors(doc) {
    var fixedCount = 0;
    
    if (!CONFIG.hyperlinkColor.applyBlack) {
        return fixedCount;
    }
    
    var blackColor;
    try {
        blackColor = doc.colors.itemByName("Black");
        if (!blackColor.isValid) {
            blackColor = doc.colors.item(0);
        }
    } catch (e) {
        blackColor = doc.colors.item(0);
    }
    
    var hyperlinks = doc.hyperlinks;
    
    for (var i = 0; i < hyperlinks.length; i++) {
        try {
            var source = hyperlinks[i].source;
            if (source) {
                source.fillColor = blackColor;
                if (CONFIG.hyperlinkColor.removeUnderline) {
                    source.underline = false;
                }
                fixedCount++;
            }
        } catch (e) {}
    }
    
    var textSources = doc.hyperlinkTextSources;
    
    for (var i = 0; i < textSources.length; i++) {
        try {
            textSources[i].fillColor = blackColor;
            if (CONFIG.hyperlinkColor.removeUnderline) {
                textSources[i].underline = false;
            }
            fixedCount++;
        } catch (e) {}
    }
    
    return fixedCount;
}

// 段落数をカウント
function countParagraphs(doc) {
    var count = 0;
    for (var i = 0; i < doc.stories.length; i++) {
        count += doc.stories[i].paragraphs.length;
    }
    return count;
}

// Wordからインポートした表を処理
function processImportedTables(doc) {
    var result = {
        tablesProcessed: 0,
        cellsProcessed: 0
    };

    if (!CONFIG.tableSettings.enabled) {
        debugLog("表処理が無効化されています");
        return result;
    }

    debugLog("=== 表（テーブル）の処理開始 ===");

    // フォント取得
    var targetFont, targetBoldFont;
    if (CONFIG.tableSettings.applyDefaultFont) {
        try {
            targetFont = app.fonts.itemByName(CONFIG.defaultFont.family + "\t" + CONFIG.defaultFont.style);
        } catch (e) {
            try {
                targetFont = app.fonts.itemByName(CONFIG.defaultFont.family);
            } catch (e2) {
                debugLog("警告: デフォルトフォントが見つかりません");
            }
        }

        try {
            targetBoldFont = app.fonts.itemByName(CONFIG.defaultFont.family + "\tBold");
        } catch (e) {
            targetBoldFont = targetFont;
        }
    }

    // 罫線の色を取得
    var strokeColor;
    try {
        strokeColor = doc.colors.itemByName(CONFIG.tableSettings.strokeColor);
        if (!strokeColor.isValid) {
            strokeColor = doc.colors.itemByName("Black");
        }
    } catch (e) {
        strokeColor = doc.colors.item(0);
    }

    // ドキュメント内の全ストーリーを検索
    var stories = doc.stories;

    for (var i = 0; i < stories.length; i++) {
        var tables = stories[i].tables;

        if (tables.length === 0) {
            continue;
        }

        debugLog("ストーリー " + i + " に " + tables.length + " 個の表を検出");

        for (var j = 0; j < tables.length; j++) {
            try {
                var table = tables[j];

                debugLog("表 " + (j + 1) + ": " + table.rows.length + "行 x " + table.columns.length + "列");

                // 表全体の罫線を設定
                try {
                    // 外枠の罫線
                    table.topBorderStrokeWeight = CONFIG.tableSettings.strokeWeight;
                    table.bottomBorderStrokeWeight = CONFIG.tableSettings.strokeWeight;
                    table.leftBorderStrokeWeight = CONFIG.tableSettings.strokeWeight;
                    table.rightBorderStrokeWeight = CONFIG.tableSettings.strokeWeight;
                    table.topBorderStrokeColor = strokeColor;
                    table.bottomBorderStrokeColor = strokeColor;
                    table.leftBorderStrokeColor = strokeColor;
                    table.rightBorderStrokeColor = strokeColor;
                } catch (e) {
                    debugLog("表罫線設定エラー: " + e.message);
                }

                // 各セルを処理
                var cells = table.cells;
                for (var k = 0; k < cells.length; k++) {
                    try {
                        var cell = cells[k];

                        // セルの内部余白を設定
                        cell.topInset = CONFIG.tableSettings.cellInset.top;
                        cell.bottomInset = CONFIG.tableSettings.cellInset.bottom;
                        cell.leftInset = CONFIG.tableSettings.cellInset.left;
                        cell.rightInset = CONFIG.tableSettings.cellInset.right;

                        // セルの罫線を設定
                        cell.topEdgeStrokeWeight = CONFIG.tableSettings.strokeWeight;
                        cell.bottomEdgeStrokeWeight = CONFIG.tableSettings.strokeWeight;
                        cell.leftEdgeStrokeWeight = CONFIG.tableSettings.strokeWeight;
                        cell.rightEdgeStrokeWeight = CONFIG.tableSettings.strokeWeight;
                        cell.topEdgeStrokeColor = strokeColor;
                        cell.bottomEdgeStrokeColor = strokeColor;
                        cell.leftEdgeStrokeColor = strokeColor;
                        cell.rightEdgeStrokeColor = strokeColor;

                        // フォントを適用
                        if (targetFont && CONFIG.tableSettings.applyDefaultFont) {
                            var isHeaderRow = (cell.rowSpan > 0 && table.rows[0] &&
                                             cell.parentRow === table.rows[0]);

                            // ヘッダー行の場合は太字、それ以外は通常
                            var cellTexts = cell.texts;
                            for (var t = 0; t < cellTexts.length; t++) {
                                try {
                                    if (isHeaderRow && CONFIG.tableSettings.headerRowBold && targetBoldFont) {
                                        cellTexts[t].appliedFont = targetBoldFont;
                                    } else {
                                        cellTexts[t].appliedFont = targetFont;
                                    }
                                    cellTexts[t].pointSize = CONFIG.fontSize;
                                } catch (e) {}
                            }
                        }

                        result.cellsProcessed++;
                    } catch (e) {
                        debugLog("セル処理エラー: " + e.message);
                    }
                }

                result.tablesProcessed++;

            } catch (e) {
                debugLog("表処理エラー: " + e.message);
            }
        }
    }

    debugLog("表処理完了: " + result.tablesProcessed + "表, " + result.cellsProcessed + "セル");
    return result;
}

// スクリプト実行
main();
