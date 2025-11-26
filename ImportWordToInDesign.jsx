// ImportWordToInDesign.jsx
// Word文書をInDesignにインポートし、段落スタイルを変換するスクリプト
// InDesign 2025対応版

#target indesign

// グローバル設定
var CONFIG = {
    autoCreatePages: true,
    maxSpreads: 100,

    // 使用するマスターページ
    masterPagePrefix: "H",
    masterPageName: "本文マスター",

    // スタイルマッピング (Word → InDesign)
    // Wordスタイル名: InDesignスタイル名
    styleMapping: {
        // 大項目 → 大見出し1
        "大項目": "大見出し1",

        // 小項目 → 小項目 (□記号を保持)
        "小項目": "小項目",

        // 標準 → Normal (インポートしたストーリーのみ)
        "標準": "Normal",

        // 演習タイトル → 演習タイトル
        "演習タイトル": "演習タイトル",

        // 図表番号 → 図番号
        "図表番号": "図番号",

        // リスト → リスト (同名でもInDesign側スタイルを明示的に適用)
        "リスト": "リスト",

        // 番号 → 番号リスト
        "番号": "番号リスト"
    },

    // 小項目の□記号設定
    kokomokuSymbol: "□　", // □+全角スペース

    // マージン設定(フォールバック用)
    fallbackMargin: {
        top: 36,
        bottom: 36,
        inside: 54,
        outside: 36
    },

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
        alert("エラー: 「" + CONFIG.masterPagePrefix + "-" + CONFIG.masterPageName + "」マスターページが見つかりません。\n\n使用可能なマスターページ:\n" + listMasterPages(doc));
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
    confirmMsg += "・大項目 → 大見出し1 (■削除)\n";
    confirmMsg += "・小項目 → 小項目 (□記号保持)\n";
    confirmMsg += "・標準/Normal → Normal\n";
    confirmMsg += "・演習タイトル → 演習タイトル\n";
    confirmMsg += "・図表番号 → 図番号\n";
    confirmMsg += "・リスト → リスト\n";
    confirmMsg += "・番号 → 番号リスト\n\n";
    confirmMsg += "【表のフォント更新】\n";
    confirmMsg += "・既存ドキュメントの全表: MS明朝 → BIZ UDゴシック\n\n";
    confirmMsg += "【マスターページ】\n";
    confirmMsg += "・" + CONFIG.masterPagePrefix + "-" + CONFIG.masterPageName + "\n\n";
    confirmMsg += "実行しますか?";

    if (!confirm(confirmMsg)) {
        return;
    }

    try {
        app.scriptPreferences.enableRedraw = false;

        var startTime = new Date();

        // 現在のページ数を記録
        var initialPageCount = doc.pages.length;

        // Word文書をインポート
        var result = importWordDocument(doc, wordFile, hMaster);

        // インポートしたストーリーのみに処理を適用
        if (result.importedStory) {
            // 段落スタイルをマッピング
            result.stylesApplied = applyStyleMapping(doc, result.importedStory);

            // 小項目に□記号を追加（まだない場合）
            result.kokomokuFixed = addKokomokuSymbol(doc, result.importedStory);

            // MS明朝 Bold → BIZ UDゴシック Regular に置換
            result.fontsReplaced = replaceFonts(result.importedStory);
        }

        // 既存ドキュメントのすべての表のフォントを更新
        result.tableFontsUpdated = updateAllTableFonts(doc);

        var endTime = new Date();
        var duration = (endTime - startTime) / 1000;

        app.scriptPreferences.enableRedraw = true;

        var resultMsg = "完了!\n\n";
        resultMsg += "既存ページ: " + initialPageCount + "p\n";
        resultMsg += "追加見開き: " + result.spreadsCreated + "\n";
        resultMsg += "段落数: " + result.paragraphsImported + "\n";
        resultMsg += "スタイル変換: " + result.stylesApplied + "件\n";
        resultMsg += "小項目□追加: " + result.kokomokuFixed + "件\n";
        resultMsg += "フォント置換: " + (result.fontsReplaced || 0) + "件\n";
        resultMsg += "表フォント更新: " + (result.tableFontsUpdated || 0) + "文字\n";
        resultMsg += "処理時間: " + duration.toFixed(1) + "秒";

        alert(resultMsg);

    } catch (e) {
        app.scriptPreferences.enableRedraw = true;
        alert("エラー:\n\n" + e.message + "\n\n行: " + e.line);
    }
}

// マスターページ一覧を取得
function listMasterPages(doc) {
    var list = [];
    for (var i = 0; i < doc.masterSpreads.length; i++) {
        list.push(doc.masterSpreads[i].name);
    }
    return list.join("\n");
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

// Word文書をインポート
function importWordDocument(doc, wordFile, master) {
    var result = {
        spreadsCreated: 0,
        paragraphsImported: 0,
        stylesApplied: 0,
        kokomokuFixed: 0,
        importedStory: null  // インポートしたストーリーを記録
    };

    // 最後のページを取得
    var lastPage = doc.pages[doc.pages.length - 1];

    // テキストフレームを取得または作成
    var textFrame = getOrCreateTextFrame(doc, lastPage);

    if (!textFrame) {
        throw new Error("テキストフレームを作成できませんでした。");
    }

    debugLog("テキストフレームを取得: " + textFrame.id);

    // Word文書を配置
    try {
        textFrame.place(wordFile);
        debugLog("Word文書を配置完了");

        // インポートしたストーリーを記録
        result.importedStory = textFrame.parentStory;
        debugLog("インポートしたストーリーID: " + result.importedStory.id);
    } catch (e) {
        throw new Error("Word文書の読み込み失敗: " + e.message);
    }

    // オーバーフローしている場合、見開きページを自動追加
    if (textFrame.overflows && CONFIG.autoCreatePages) {
        debugLog("テキストがオーバーフロー、ページを追加します");
        result.spreadsCreated = autoFlowWithSpreads(doc, textFrame, master);
    }

    result.paragraphsImported = result.importedStory ? result.importedStory.paragraphs.length : 0;

    return result;
}

// テキストフレームを取得または作成
function getOrCreateTextFrame(doc, page) {
    // まずマスターページのテキストフレームをオーバーライド
    var masterFrame = overrideMasterTextFrame(page);
    if (masterFrame) {
        debugLog("マスターフレームをオーバーライド");
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
        debugLog("既存のテキストフレームを使用");
        return largestFrame;
    }

    // テキストフレームがない場合は新規作成
    debugLog("新規テキストフレームを作成");
    return createTextFrame(doc, page);
}

// マスターページのテキストフレームをオーバーライド
function overrideMasterTextFrame(page) {
    try {
        var appliedMaster = page.appliedMaster;

        if (!appliedMaster || appliedMaster.name === "None") {
            return null;
        }

        // ページ上のすべてのマスターアイテムをオーバーライド
        try {
            page.overrideAll();

            var textFrames = page.textFrames;

            // 連結されていない最初のテキストフレームを返す
            for (var i = 0; i < textFrames.length; i++) {
                if (!textFrames[i].previousTextFrame) {
                    return textFrames[i];
                }
            }

            if (textFrames.length > 0) {
                return textFrames[0];
            }
        } catch (e) {
            debugLog("overrideAll失敗: " + e.message);
        }

        // 手動でオーバーライド
        var masterPages = appliedMaster.pages;

        for (var i = 0; i < masterPages.length; i++) {
            var masterPage = masterPages[i];

            if (masterPage.side === page.side || masterPage.side === PageSideOptions.SINGLE_SIDED) {
                var masterTextFrames = masterPage.textFrames;

                if (masterTextFrames.length > 0) {
                    try {
                        return masterTextFrames[0].override(page);
                    } catch (e) {
                        debugLog("個別オーバーライド失敗: " + e.message);
                    }
                }
            }
        }

    } catch (e) {
        debugLog("マスターオーバーライドエラー: " + e.message);
    }

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

        // 新しい見開きを追加
        var newSpread = doc.spreads.add(LocationOptions.AT_END);
        spreadsCreated++;

        debugLog("見開き " + spreadsCreated + " を追加");

        // マスターページを適用
        try {
            newSpread.appliedMaster = master;
        } catch (e) {
            debugLog("マスター適用失敗: " + e.message);
        }

        var pages = newSpread.pages;

        if (pages.length === 0) {
            break;
        }

        // 見開きの各ページにテキストフレームを連結
        for (var i = 0; i < pages.length; i++) {
            var pageFrame = getOrCreateTextFrame(doc, pages[i]);

            if (pageFrame) {
                try {
                    // 前のフレームと連結
                    pageFrame.previousTextFrame = currentFrame;
                    currentFrame = pageFrame;

                    debugLog("ページ " + pages[i].name + " にフレームを連結");

                    // オーバーフローがなくなったら終了
                    if (!currentFrame.overflows) {
                        debugLog("オーバーフロー解消");
                        break;
                    }
                } catch (e) {
                    debugLog("フレーム連結エラー: " + e.message);
                    return spreadsCreated;
                }
            }
        }

        // 最後のフレームに移動
        while (currentFrame.nextTextFrame) {
            currentFrame = currentFrame.nextTextFrame;
        }
    }

    debugLog("合計 " + spreadsCreated + " 見開きを追加");
    return spreadsCreated;
}

// スタイルマッピングを適用（インポートしたストーリーのみ）
function applyStyleMapping(doc, importedStory) {
    var mappingCount = 0;

    debugLog("=== スタイルマッピング開始（インポートしたストーリーのみ） ===");

    if (!importedStory || !importedStory.isValid) {
        debugLog("有効なストーリーがありません");
        return 0;
    }

    var paragraphs = importedStory.paragraphs;

    for (var j = 0; j < paragraphs.length; j++) {
        var para = paragraphs[j];
        var currentStyleName = para.appliedParagraphStyle.name;

        // スタイルマッピングに該当するか確認
        if (CONFIG.styleMapping.hasOwnProperty(currentStyleName)) {
            var targetStyleName = CONFIG.styleMapping[currentStyleName];

            try {
                var targetStyle = doc.paragraphStyles.itemByName(targetStyleName);
                if (targetStyle.isValid) {
                    // 大項目→大見出し1の場合、先頭の■を削除
                    if (currentStyleName === "大項目" && targetStyleName === "大見出し1") {
                        removeLeadingSymbol(para, "■");
                    }

                    para.appliedParagraphStyle = targetStyle;
                    mappingCount++;

                    if (mappingCount <= 5) {
                        debugLog("スタイル変換: " + currentStyleName + " → " + targetStyleName);
                    }
                } else {
                    debugLog("対象スタイルが見つかりません: " + targetStyleName);
                }
            } catch (e) {
                debugLog("スタイル適用エラー: " + e.message);
            }
        }
    }

    debugLog("スタイルマッピング完了: " + mappingCount + "件");
    return mappingCount;
}

// 段落先頭の記号を削除
function removeLeadingSymbol(para, symbol) {
    try {
        var paraText = para.contents;

        // 先頭が指定記号で始まっている場合
        if (paraText.indexOf(symbol) === 0) {
            // 記号とその後のスペース（全角・半角）を削除
            var newText = paraText.substring(1);
            // 先頭のスペースも削除
            newText = newText.replace(/^[\s　]+/, "");
            para.contents = newText;
            debugLog("■記号を削除: " + symbol);
        }
    } catch (e) {
        debugLog("記号削除エラー: " + e.message);
    }
}

// 小項目に□記号を追加（インポートしたストーリーのみ）
function addKokomokuSymbol(doc, importedStory) {
    var addCount = 0;
    var kokomokuStyle = doc.paragraphStyles.itemByName("小項目");

    if (!kokomokuStyle.isValid) {
        debugLog("「小項目」スタイルが見つかりません");
        return 0;
    }

    if (!importedStory || !importedStory.isValid) {
        debugLog("有効なストーリーがありません");
        return 0;
    }

    debugLog("=== 小項目に□記号を追加開始（インポートしたストーリーのみ） ===");

    var symbol = CONFIG.kokomokuSymbol;
    var paragraphs = importedStory.paragraphs;

    for (var j = 0; j < paragraphs.length; j++) {
        var para = paragraphs[j];

        if (para.appliedParagraphStyle.name === "小項目") {
            try {
                var paraText = para.contents;

                // 既に□で始まっている場合はスキップ
                if (paraText.indexOf("□") === 0) {
                    debugLog("既に□あり: 段落 " + j);
                    continue;
                }

                // 段落の先頭に□記号を挿入
                para.contents = symbol + paraText;
                addCount++;

                if (addCount <= 5) {
                    debugLog("□記号追加: 段落 " + j);
                }
            } catch (e) {
                debugLog("□記号追加エラー: " + e.message);
            }
        }
    }

    debugLog("小項目□記号追加完了: " + addCount + "件");
    return addCount;
}

// フォント置換（インポートしたストーリーのみ）
// MS明朝 Bold → BIZ UDゴシック Regular
function replaceFonts(importedStory) {
    var replaceCount = 0;

    if (!importedStory || !importedStory.isValid) {
        debugLog("有効なストーリーがありません");
        return 0;
    }

    debugLog("=== フォント置換開始（インポートしたストーリーのみ） ===");

    // 置換先フォントを取得
    var targetFont;
    try {
        targetFont = app.fonts.itemByName("BIZ UDGothic\tRegular");
        if (!targetFont.isValid) {
            targetFont = app.fonts.itemByName("BIZ UDゴシック\tRegular");
        }
    } catch (e) {
        debugLog("BIZ UDゴシックが見つかりません: " + e.message);
        return 0;
    }

    if (!targetFont || !targetFont.isValid) {
        debugLog("置換先フォントが見つかりません");
        return 0;
    }

    debugLog("置換先フォント: " + targetFont.name);

    // テキストを走査してフォントを置換
    try {
        var characters = importedStory.characters;

        for (var i = 0; i < characters.length; i++) {
            try {
                var ch = characters[i];
                var fontName = ch.appliedFont.name;

                // MS明朝 Bold を検出（様々な表記に対応）
                if (fontName.indexOf("MS") >= 0 && fontName.indexOf("明朝") >= 0 && fontName.indexOf("Bold") >= 0) {
                    ch.appliedFont = targetFont;
                    replaceCount++;
                }
                // ＭＳ 明朝 Bold（全角）
                else if (fontName.indexOf("ＭＳ") >= 0 && fontName.indexOf("明朝") >= 0 && fontName.indexOf("Bold") >= 0) {
                    ch.appliedFont = targetFont;
                    replaceCount++;
                }
            } catch (e) {
                // 個別の文字エラーは無視
            }
        }
    } catch (e) {
        debugLog("フォント置換エラー: " + e.message);
    }

    debugLog("フォント置換完了: " + replaceCount + "件");
    return replaceCount;
}

// 既存ドキュメントのすべての表のフォントを更新
// MS明朝（Regular/Bold問わず）→ BIZ UDゴシック Regular
function updateAllTableFonts(doc) {
    var replaceCount = 0;
    var tableCount = 0;

    debugLog("=== 既存ドキュメントの全表フォント更新開始 ===");

    // 置換先フォントを取得
    var targetFont;
    try {
        targetFont = app.fonts.itemByName("BIZ UDGothic\tRegular");
        if (!targetFont.isValid) {
            targetFont = app.fonts.itemByName("BIZ UDゴシック\tRegular");
        }
    } catch (e) {
        debugLog("BIZ UDゴシックが見つかりません: " + e.message);
        return 0;
    }

    if (!targetFont || !targetFont.isValid) {
        debugLog("置換先フォントが見つかりません");
        return 0;
    }

    debugLog("置換先フォント: " + targetFont.name);

    // ドキュメント内のすべてのストーリーを走査
    for (var s = 0; s < doc.stories.length; s++) {
        var story = doc.stories[s];

        // ストーリー内のすべてのテーブルを走査
        var tables = story.tables;
        for (var t = 0; t < tables.length; t++) {
            var table = tables[t];
            tableCount++;

            debugLog("表 " + tableCount + " を処理中...");

            // テーブル内のすべてのセルを走査
            var cells = table.cells;
            for (var c = 0; c < cells.length; c++) {
                var cell = cells[c];

                // セル内のすべての文字を走査
                try {
                    var characters = cell.characters;
                    for (var i = 0; i < characters.length; i++) {
                        try {
                            var ch = characters[i];
                            var fontName = ch.appliedFont.name;

                            // MS明朝を検出（Regular/Bold問わず）
                            var isMincho = false;
                            if (fontName.indexOf("MS") >= 0 && fontName.indexOf("明朝") >= 0) {
                                isMincho = true;
                            } else if (fontName.indexOf("ＭＳ") >= 0 && fontName.indexOf("明朝") >= 0) {
                                isMincho = true;
                            } else if (fontName.indexOf("Mincho") >= 0) {
                                isMincho = true;
                            }

                            if (isMincho) {
                                ch.appliedFont = targetFont;
                                replaceCount++;
                            }
                        } catch (e) {
                            // 個別の文字エラーは無視
                        }
                    }
                } catch (e) {
                    debugLog("セル処理エラー: " + e.message);
                }
            }
        }
    }

    debugLog("全表フォント更新完了: " + tableCount + "表, " + replaceCount + "文字");
    return replaceCount;
}

// 段落数をカウント
function countParagraphs(doc) {
    var count = 0;
    for (var i = 0; i < doc.stories.length; i++) {
        count += doc.stories[i].paragraphs.length;
    }
    return count;
}

// スクリプト実行
main();
