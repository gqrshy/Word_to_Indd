// ImportWordToInDesign.jsx
// Word文書をInDesignにインポートし、段落スタイルを変換するスクリプト
// InDesign 2025対応版

#target indesign

// ============================================================
// 設定
// ============================================================

var CONFIG = {
    autoCreatePages: true,
    maxSpreads: 100,
    masterPagePrefix: "H",
    masterPageName: "本文マスター",

    styleMapping: {
        "大項目": "大見出し1",
        "小項目": "小項目",
        "標準": "Normal",
        "演習タイトル": "演習タイトル",
        "図表番号": "図番号",
        "リスト": "リスト",
        "番号": "番号リスト"
    },

    kokomokuSymbol: "□　",

    fallbackMargin: {
        top: 36,
        bottom: 36,
        inside: 54,
        outside: 36
    },

    debugMode: true
};

function debugLog(message) {
    if (CONFIG.debugMode) {
        $.writeln("[DEBUG] " + message);
    }
}

// ============================================================
// Wordインポート
// ============================================================

function configureWordImportPreferences() {
    try {
        var prefs = app.wordRTFImportPreferences;
        prefs.removeFormatting = false;
        prefs.preserveGraphics = true;
        prefs.preserveLocalOverrides = true;
        prefs.importFootnotes = true;
        prefs.importEndnotes = true;
        prefs.importIndex = true;
        prefs.importTOC = true;
        prefs.importUnusedStyles = false;
        prefs.useTypographersQuotes = true;
        debugLog("Wordインポート設定を構成完了");
    } catch (e) {
        debugLog("Wordインポート設定エラー: " + e.message);
    }
}

function detectEmbeddedObjectIssues(story) {
    if (!story || !story.isValid) return 0;

    var issueCount = 0;
    var paragraphs = story.paragraphs;

    for (var i = 0; i < paragraphs.length; i++) {
        try {
            var para = paragraphs[i];
            var text = para.contents;
            var cleanText = text.replace(/[\r\n]/g, "");

            if (cleanText === "□" || cleanText === "\uFFFD" || cleanText === "\u25A1") {
                issueCount++;
                debugLog("警告: 段落 " + (i + 1) + " に孤立した□文字を検出");
            }
        } catch (e) {}
    }

    if (issueCount > 0) {
        debugLog("埋め込みオブジェクト問題: " + issueCount + "件");
    }

    return issueCount;
}

function importWordDocument(doc, wordFile, master) {
    var result = {
        spreadsCreated: 0,
        paragraphsImported: 0,
        stylesApplied: 0,
        kokomokuFixed: 0,
        importedStory: null
    };

    var lastPage = doc.pages[doc.pages.length - 1];
    var textFrame = getOrCreateTextFrame(doc, lastPage);

    if (!textFrame) {
        throw new Error("テキストフレームを作成できませんでした。");
    }

    debugLog("テキストフレームを取得: " + textFrame.id);

    try {
        configureWordImportPreferences();
        textFrame.place(wordFile);
        debugLog("Word文書を配置完了");

        result.importedStory = textFrame.parentStory;
        debugLog("インポートしたストーリーID: " + result.importedStory.id);

        var tableCount = result.importedStory.tables.length;
        debugLog("インポートされた表の数: " + tableCount);

        result.embeddedObjectIssues = detectEmbeddedObjectIssues(result.importedStory);
    } catch (e) {
        throw new Error("Word文書の読み込み失敗: " + e.message);
    }

    if (textFrame.overflows && CONFIG.autoCreatePages) {
        debugLog("テキストがオーバーフロー、ページを追加します");
        result.spreadsCreated = autoFlowWithSpreads(doc, textFrame, master);
    }

    result.paragraphsImported = result.importedStory ? result.importedStory.paragraphs.length : 0;

    return result;
}

// ============================================================
// テキストフレーム管理
// ============================================================

function getOrCreateTextFrame(doc, page) {
    var masterFrame = overrideMasterTextFrame(page);
    if (masterFrame) {
        debugLog("マスターフレームをオーバーライド");
        return masterFrame;
    }

    var textFrames = page.textFrames;
    var largestFrame = null;
    var largestArea = 0;

    for (var i = 0; i < textFrames.length; i++) {
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

    debugLog("新規テキストフレームを作成");
    return createTextFrame(doc, page);
}

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
        } catch (e) {
            debugLog("overrideAll失敗: " + e.message);
        }

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

function autoFlowWithSpreads(doc, startFrame, master) {
    var spreadsCreated = 0;
    var currentFrame = startFrame;
    var loopCount = 0;

    while (currentFrame.overflows && loopCount < CONFIG.maxSpreads) {
        loopCount++;

        var newSpread = doc.spreads.add(LocationOptions.AT_END);
        spreadsCreated++;

        debugLog("見開き " + spreadsCreated + " を追加");

        try {
            newSpread.appliedMaster = master;
        } catch (e) {
            debugLog("マスター適用失敗: " + e.message);
        }

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

                    debugLog("ページ " + pages[i].name + " にフレームを連結");

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

        while (currentFrame.nextTextFrame) {
            currentFrame = currentFrame.nextTextFrame;
        }
    }

    debugLog("合計 " + spreadsCreated + " 見開きを追加");
    return spreadsCreated;
}

// ============================================================
// スタイル変換
// ============================================================

function applyStyleMapping(doc, importedStory) {
    var mappingCount = 0;

    debugLog("=== スタイルマッピング開始 ===");

    if (!importedStory || !importedStory.isValid) {
        debugLog("有効なストーリーがありません");
        return 0;
    }

    var paragraphs = importedStory.paragraphs;

    for (var j = 0; j < paragraphs.length; j++) {
        var para = paragraphs[j];
        var currentStyleName = para.appliedParagraphStyle.name;

        if (CONFIG.styleMapping.hasOwnProperty(currentStyleName)) {
            var targetStyleName = CONFIG.styleMapping[currentStyleName];

            try {
                var targetStyle = doc.paragraphStyles.itemByName(targetStyleName);
                if (targetStyle.isValid) {
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

function removeLeadingSymbol(para, symbol) {
    try {
        var paraText = para.contents;

        if (paraText.indexOf(symbol) === 0) {
            var newText = paraText.substring(1);
            newText = newText.replace(/^[\s　]+/, "");
            para.contents = newText;
            debugLog("■記号を削除");
        }
    } catch (e) {
        debugLog("記号削除エラー: " + e.message);
    }
}

function addKokomokuSymbol(doc, importedStory) {
    var addCount = 0;
    var skippedCount = 0;
    var kokomokuStyle = doc.paragraphStyles.itemByName("小項目");

    if (!kokomokuStyle.isValid) {
        debugLog("「小項目」スタイルが見つかりません");
        return 0;
    }

    if (!importedStory || !importedStory.isValid) {
        debugLog("有効なストーリーがありません");
        return 0;
    }

    debugLog("=== 小項目に□記号を追加開始 ===");

    var symbol = CONFIG.kokomokuSymbol;
    var paragraphs = importedStory.paragraphs;

    for (var j = 0; j < paragraphs.length; j++) {
        var para = paragraphs[j];

        if (para.appliedParagraphStyle.name === "小項目") {
            try {
                var paraText = para.contents;

                if (paraText.indexOf("□") === 0) {
                    debugLog("既に□あり: 段落 " + j);
                    continue;
                }

                var hasInlineObjects = false;

                if (para.tables && para.tables.length > 0) {
                    hasInlineObjects = true;
                    debugLog("警告: 段落 " + j + " に表が含まれています - スキップ");
                }

                if (!hasInlineObjects) {
                    for (var c = 0; c < paraText.length; c++) {
                        var charCode = paraText.charCodeAt(c);
                        if (charCode === 0xFFFC || charCode < 0x0020 && charCode !== 0x0009 && charCode !== 0x000A && charCode !== 0x000D) {
                            hasInlineObjects = true;
                            debugLog("警告: 段落 " + j + " にインラインオブジェクトを検出 - スキップ");
                            break;
                        }
                    }
                }

                var trimmedText = paraText.replace(/[\s\r\n　]/g, "");
                if (trimmedText.length === 0) {
                    skippedCount++;
                    debugLog("空の小項目段落をスキップ: 段落 " + j);
                    continue;
                }

                if (hasInlineObjects) {
                    skippedCount++;
                    continue;
                }

                try {
                    para.insertionPoints[0].contents = symbol;
                    addCount++;
                } catch (insertError) {
                    para.contents = symbol + paraText;
                    addCount++;
                }

                if (addCount <= 5) {
                    debugLog("□記号追加: 段落 " + j);
                }
            } catch (e) {
                debugLog("□記号追加エラー: " + e.message);
            }
        }
    }

    debugLog("小項目□記号追加完了: " + addCount + "件");
    if (skippedCount > 0) {
        debugLog("スキップした段落: " + skippedCount + "件");
    }
    return addCount;
}

// ============================================================
// フォント置換
// ============================================================

function replaceFonts(importedStory) {
    var replaceCount = 0;

    if (!importedStory || !importedStory.isValid) {
        debugLog("有効なストーリーがありません");
        return 0;
    }

    debugLog("=== フォント置換開始 ===");

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

    try {
        var characters = importedStory.characters;

        for (var i = 0; i < characters.length; i++) {
            try {
                var ch = characters[i];
                var fontName = ch.appliedFont.name;

                if (fontName.indexOf("MS") >= 0 && fontName.indexOf("明朝") >= 0 && fontName.indexOf("Bold") >= 0) {
                    ch.appliedFont = targetFont;
                    replaceCount++;
                } else if (fontName.indexOf("ＭＳ") >= 0 && fontName.indexOf("明朝") >= 0 && fontName.indexOf("Bold") >= 0) {
                    ch.appliedFont = targetFont;
                    replaceCount++;
                }
            } catch (e) {}
        }
    } catch (e) {
        debugLog("フォント置換エラー: " + e.message);
    }

    debugLog("フォント置換完了: " + replaceCount + "件");
    return replaceCount;
}

function updateAllTableFonts(doc) {
    var replaceCount = 0;
    var tableCount = 0;

    debugLog("=== 全表フォント更新開始 ===");

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

    for (var s = 0; s < doc.stories.length; s++) {
        var story = doc.stories[s];

        var tables = story.tables;
        for (var t = 0; t < tables.length; t++) {
            var table = tables[t];
            tableCount++;

            debugLog("表 " + tableCount + " を処理中...");

            var cells = table.cells;
            for (var c = 0; c < cells.length; c++) {
                var cell = cells[c];

                try {
                    var characters = cell.characters;
                    for (var i = 0; i < characters.length; i++) {
                        try {
                            var ch = characters[i];
                            var fontName = ch.appliedFont.name;

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
                        } catch (e) {}
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

// ============================================================
// アンカードオブジェクト処理
// ============================================================

function applyCodeStyleToAnchoredObjects(doc, importedStory) {
    var styleApplied = 0;

    debugLog("=== アンカードオブジェクト内テキストにスタイル適用開始 ===");

    if (!importedStory || !importedStory.isValid) {
        debugLog("有効なストーリーがありません");
        return 0;
    }

    var codeStyle;
    try {
        codeStyle = doc.paragraphStyles.itemByName("コード・コマンド");
        if (!codeStyle.isValid) {
            debugLog("「コード・コマンド」スタイルが見つかりません");
            return 0;
        }
    } catch (e) {
        debugLog("「コード・コマンド」スタイル取得エラー: " + e.message);
        return 0;
    }

    debugLog("「コード・コマンド」スタイルを使用します");

    var textContainers = importedStory.textContainers;

    for (var i = 0; i < textContainers.length; i++) {
        var container = textContainers[i];

        if (container.constructor.name !== "TextFrame") {
            continue;
        }

        try {
            var allPageItems = container.allPageItems;

            for (var j = 0; j < allPageItems.length; j++) {
                var item = allPageItems[j];
                styleApplied += processAnchoredItem(item, codeStyle);
            }
        } catch (e) {
            debugLog("アンカードオブジェクト検索エラー: " + e.message);
        }
    }

    try {
        var paragraphs = importedStory.paragraphs;
        for (var p = 0; p < paragraphs.length; p++) {
            try {
                var para = paragraphs[p];
                if (para.allPageItems && para.allPageItems.length > 0) {
                    for (var k = 0; k < para.allPageItems.length; k++) {
                        styleApplied += processAnchoredItem(para.allPageItems[k], codeStyle);
                    }
                }
            } catch (e) {}
        }
    } catch (e) {
        debugLog("段落内アイテム処理エラー: " + e.message);
    }

    debugLog("アンカードオブジェクト内テキストスタイル適用完了: " + styleApplied + "段落");
    return styleApplied;
}

function processAnchoredItem(item, codeStyle) {
    var styleApplied = 0;

    try {
        if (item.constructor.name === "Group") {
            var groupItems = item.allPageItems;
            for (var i = 0; i < groupItems.length; i++) {
                styleApplied += processAnchoredItem(groupItems[i], codeStyle);
            }
        } else if (item.constructor.name === "TextFrame") {
            var textFrame = item;

            try {
                var paragraphs = textFrame.paragraphs;

                for (var p = 0; p < paragraphs.length; p++) {
                    try {
                        var para = paragraphs[p];
                        var content = para.contents.replace(/[\r\n\s　]/g, "");
                        if (content.length === 0) {
                            continue;
                        }

                        para.appliedParagraphStyle = codeStyle;
                        styleApplied++;

                        if (styleApplied <= 5) {
                            debugLog("図内テキストにスタイル適用: " + para.contents.substring(0, 30) + "...");
                        }
                    } catch (e) {}
                }
            } catch (e) {
                debugLog("テキストフレーム処理エラー: " + e.message);
            }
        } else if (item.contentType === ContentType.TEXT_TYPE) {
            try {
                var paragraphs = item.paragraphs;
                for (var p = 0; p < paragraphs.length; p++) {
                    try {
                        var para = paragraphs[p];
                        var content = para.contents.replace(/[\r\n\s　]/g, "");
                        if (content.length === 0) {
                            continue;
                        }
                        para.appliedParagraphStyle = codeStyle;
                        styleApplied++;
                    } catch (e) {}
                }
            } catch (e) {}
        }
    } catch (e) {}

    return styleApplied;
}

// ============================================================
// ユーティリティ
// ============================================================

function listMasterPages(doc) {
    var list = [];
    for (var i = 0; i < doc.masterSpreads.length; i++) {
        list.push(doc.masterSpreads[i].name);
    }
    return list.join("\n");
}

function getMasterPage(doc, masterName) {
    try {
        var master = doc.masterSpreads.itemByName(masterName);
        if (master.isValid) {
            return master;
        }
    } catch (e) {}
    return null;
}

function countParagraphs(doc) {
    var count = 0;
    for (var i = 0; i < doc.stories.length; i++) {
        count += doc.stories[i].paragraphs.length;
    }
    return count;
}

// ============================================================
// メイン処理
// ============================================================

function main() {
    if (app.documents.length === 0) {
        alert("InDesignドキュメントを開いてからスクリプトを実行してください。");
        return;
    }

    var doc = app.activeDocument;

    var hMaster = getMasterPage(doc, CONFIG.masterPagePrefix + "-" + CONFIG.masterPageName);
    if (!hMaster) {
        alert("エラー: 「" + CONFIG.masterPagePrefix + "-" + CONFIG.masterPageName + "」マスターページが見つかりません。\n\n使用可能なマスターページ:\n" + listMasterPages(doc));
        return;
    }

    debugLog("使用マスター: " + hMaster.name);

    var wordFile = File.openDialog("インポートするWord文書を選択してください (.docx)", "*.docx");
    if (!wordFile) {
        return;
    }

    var confirmMsg = "Word文書をインポートします\n\n";
    confirmMsg += "【段落スタイル変換】\n";
    confirmMsg += "・大項目 → 大見出し1 (■削除)\n";
    confirmMsg += "・小項目 → 小項目 (□記号保持)\n";
    confirmMsg += "・標準/Normal → Normal\n";
    confirmMsg += "・演習タイトル → 演習タイトル\n";
    confirmMsg += "・図表番号 → 図番号\n";
    confirmMsg += "・リスト → リスト\n";
    confirmMsg += "・番号 → 番号リスト\n";
    confirmMsg += "・図内テキスト → コード・コマンド\n\n";
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
        var initialPageCount = doc.pages.length;

        var result = importWordDocument(doc, wordFile, hMaster);

        if (result.importedStory) {
            result.stylesApplied = applyStyleMapping(doc, result.importedStory);
            result.kokomokuFixed = addKokomokuSymbol(doc, result.importedStory);
            result.fontsReplaced = replaceFonts(result.importedStory);
            result.codeStyleApplied = applyCodeStyleToAnchoredObjects(doc, result.importedStory);
        }

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
        resultMsg += "図内コード・コマンド: " + (result.codeStyleApplied || 0) + "段落\n";
        resultMsg += "フォント置換: " + (result.fontsReplaced || 0) + "件\n";
        resultMsg += "表フォント更新: " + (result.tableFontsUpdated || 0) + "文字\n";
        resultMsg += "処理時間: " + duration.toFixed(1) + "秒";

        if (result.embeddedObjectIssues && result.embeddedObjectIssues > 0) {
            resultMsg += "\n\n⚠ 警告: " + result.embeddedObjectIssues + "件の埋め込みオブジェクト問題を検出\n";
            resultMsg += "（□のみの段落がある場合、Wordの埋め込み表が\n";
            resultMsg += "変換できなかった可能性があります。\n";
            resultMsg += "Wordで表を選択→右クリック→「表に変換」後、\n";
            resultMsg += "再度インポートしてください）";
        }

        alert(resultMsg);

    } catch (e) {
        app.scriptPreferences.enableRedraw = true;
        alert("エラー:\n\n" + e.message + "\n\n行: " + e.line);
    }
}

main();
