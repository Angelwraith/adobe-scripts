/*@METADATA{
  "name": "List Artboard Sizes",
  "description": "Lists all artboards across one or more open documents with sizes multiplied by 10 and rounded up to the nearest 0.5 inch. Shows trim size plus finished (with-bleed) size for each panel. Copy for Email can include the panel size table as a clean HTML table. Also copies shape duplicates. Supports a custom (unusual) roll width in addition to the stocked rolls. Files with PRIME in the name are treated as nested sheets (bleed already included) and compared against the panel files, with a waste breakdown (bleed / unused roll width / R.L. waste).",
  "version": "1.8",
  "target": "illustrator",
  "tags": ["artboard", "size", "measure", "list", "multi-document"]
}@END_METADATA*/

#target illustrator

// ============================================================================
// CONSTANTS (must be defined before any function that runs at startup)
// ============================================================================
// Defaults match the production spreadsheet for Promaster / wrap jobs.
// All values in inches. Roll widths are nominal; pinchMargin is reserved per
// roll (not usable), so usable width = rollWidth - pinchMargin.
var ESTIMATION_DEFAULTS = {
    bleed: 2,                       // inches per side
    rlWaste: 8,                     // inches added to run length per panel
    pinchMargin: 2,                 // inches lost to pinch rollers per roll
    rollWidths: [36, 48, 54, 60]    // available stocked rolls
};

// ============================================================================
// MAIN ENTRY POINT
// ============================================================================
try {
    if (app.documents.length === 0) {
        alert("Please open at least one document first.");
    } else {
        listArtboardSizes();
    }
} catch (e) {
    alert("Startup error: " + e.toString());
}

// ============================================================================
// MAIN FLOW
// ============================================================================
function listArtboardSizes() {
    var selectedDocs = showDocumentSelectionDialog();
    if (!selectedDocs || selectedDocs.length === 0) return;

    var multiDoc = selectedDocs.length > 1;
    var rows = collectArtboardRows(selectedDocs, multiDoc);

    if (rows.length === 0) {
        alert("No artboards to list (all skipped or named 'proof').");
        return;
    }

    showResults(rows, multiDoc);
}

// ============================================================================
// DOCUMENT SELECTION DIALOG
// ============================================================================
function showDocumentSelectionDialog() {
    // Filter out any documents with "proof" in the name -- never useful for sizing
    var eligibleDocs = [];
    for (var d = 0; d < app.documents.length; d++) {
        if (app.documents[d].name.toLowerCase().indexOf("proof") === -1) {
            eligibleDocs.push(app.documents[d]);
        }
    }

    if (eligibleDocs.length === 0) {
        alert("No eligible documents are open (all open files have 'proof' in the name).");
        return null;
    }

    var dialog = new Window("dialog", "List Artboard Sizes - Select Documents");
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.spacing = 15;
    dialog.margins = 20;
    dialog.preferredSize.width = 480;

    var titleText = dialog.add("statictext", undefined, "Select documents to scan:");
    titleText.graphics.font = ScriptUI.newFont("dialog", "Bold", 12);

    var listPanel = dialog.add("panel");
    listPanel.orientation = "column";
    listPanel.alignChildren = "fill";
    listPanel.margins = 10;
    listPanel.preferredSize.height = 200;

    var activeDocName = app.activeDocument.name;
    var checkboxes = [];
    for (var i = 0; i < eligibleDocs.length; i++) {
        var doc = eligibleDocs[i];
        var cb = listPanel.add("checkbox", undefined, doc.name);
        cb.value = (doc.name === activeDocName);
        cb.preferredSize.width = 420;
        checkboxes.push({ checkbox: cb, document: doc });
    }

    var selectionGroup = dialog.add("group");
    selectionGroup.alignment = "center";
    selectionGroup.spacing = 10;

    var selectAllBtn = selectionGroup.add("button", undefined, "Select All");
    selectAllBtn.preferredSize.width = 90;
    var selectCurrentBtn = selectionGroup.add("button", undefined, "Select Current");
    selectCurrentBtn.preferredSize.width = 100;
    var selectNoneBtn = selectionGroup.add("button", undefined, "Select None");
    selectNoneBtn.preferredSize.width = 90;

    selectAllBtn.onClick = function() {
        for (var i = 0; i < checkboxes.length; i++) checkboxes[i].checkbox.value = true;
    };
    selectCurrentBtn.onClick = function() {
        var name = app.activeDocument.name;
        for (var i = 0; i < checkboxes.length; i++) {
            checkboxes[i].checkbox.value = (checkboxes[i].document.name === name);
        }
    };
    selectNoneBtn.onClick = function() {
        for (var i = 0; i < checkboxes.length; i++) checkboxes[i].checkbox.value = false;
    };

    var btnGroup = dialog.add("group");
    btnGroup.alignment = "center";
    btnGroup.spacing = 10;
    var cancelBtn = btnGroup.add("button", undefined, "Cancel", {name: "cancel"});
    var okBtn = btnGroup.add("button", undefined, "Process", {name: "ok"});

    var result = null;
    cancelBtn.onClick = function() { result = null; dialog.close(); };
    okBtn.onClick = function() {
        var picked = [];
        for (var i = 0; i < checkboxes.length; i++) {
            if (checkboxes[i].checkbox.value) picked.push(checkboxes[i].document);
        }
        if (picked.length === 0) {
            alert("Please select at least one document.");
            return;
        }
        result = picked;
        dialog.close();
    };

    dialog.show();
    return result;
}

// ============================================================================
// COLLECT ARTBOARD DATA FROM SELECTED DOCUMENTS
// ============================================================================
function collectArtboardRows(docs, multiDoc) {
    var rows = [];

    // Remember the originally active doc so we can restore it after the scan
    var originalActive = null;
    try { originalActive = app.activeDocument; } catch (e) {}

    for (var d = 0; d < docs.length; d++) {
        var doc = docs[d];

        // CRITICAL: Illustrator's artboards collection reflects the active
        // document rather than the doc reference, so we must activate each
        // doc before reading its artboards.
        try {
            app.activeDocument = doc;
            app.redraw();
            $.sleep(50);
        } catch (e) {
            continue; // doc was closed or otherwise invalid
        }

        var baseName = doc.name.replace(/\.[^\.]+$/, "");
        // PRIME files hold nested print sheets: bleed is already in the artboards
        var isPrime = /PRIME/i.test(doc.name);

        for (var i = 0; i < doc.artboards.length; i++) {
            var ab = doc.artboards[i];

            // Skip artboards named "proof" (case-insensitive)
            if (ab.name.toLowerCase() === "proof") continue;

            var rect = ab.artboardRect; // [left, top, right, bottom] in points

            var widthPts = Math.abs(rect[2] - rect[0]);
            var heightPts = Math.abs(rect[1] - rect[3]);

            var widthIn = widthPts / 72;
            var heightIn = heightPts / 72;

            // x10, then round UP to nearest 0.5"
            var widthRounded = Math.ceil(widthIn * 10 * 2) / 2;
            var heightRounded = Math.ceil(heightIn * 10 * 2) / 2;

            // When multi-doc, prefix with doc base name to distinguish source
            var displayName = multiDoc ? (baseName + " - " + ab.name) : ab.name;

            rows.push({
                index: rows.length + 1,
                name: displayName,
                docName: doc.name,
                docBaseName: baseName,
                artboardName: ab.name,
                origRect: rect,
                origWpts: widthPts,
                origHpts: heightPts,
                finalW: widthRounded,
                finalH: heightRounded,
                isPrime: isPrime
            });
        }
    }

    // Restore original active document
    if (originalActive) {
        try { app.activeDocument = originalActive; } catch (e) {}
    }

    return rows;
}

// ============================================================================
// FORMATTING HELPERS
// ============================================================================
function formatSize(n) {
    var rounded = Math.round(n * 1000) / 1000;
    return rounded.toString();
}

// Finished (printed) dimension = trim + bleed on both sides, rounded up to 0.5".
// bleed is in the same x10 scale as the trim size (finalW/finalH).
function finishedDim(trim, bleed) {
    return Math.ceil((trim + bleed * 2) * 2) / 2;
}

function padRight(str, len) {
    var out = str;
    while (out.length < len) out += " ";
    return out;
}

function padLeft(str, len) {
    var out = str;
    while (out.length < len) out = " " + out;
    return out;
}

// Build structured rows for the multi-column list control. Returns an array of
// objects: { sep:true } marks a blank separator between document sets, otherwise
// { sep:false, num, name, trim, bleed, roll, run } holds one cell per column.
// The trim / w-bleed size pairs keep their numbers padded to fixed sub-widths so
// the "x" reads cleanly even in a proportional font.
//
// Columns:  #   Panel   Trim   W/ Bleed   Roll   Run
//   Roll = smallest stocked roll the panel fits on (from r._estimate)
//   Run  = run length consumed on that roll, incl. R.L. waste (runWithWaste)
function buildDisplayRows(rows, bleed) {
    // First pass: find sub-field widths so the size numbers line up
    var wTw = 0, wTh = 0, wBw = 0, wBh = 0;
    var rawTw = [], rawTh = [], rawBw = [], rawBh = [];
    for (var i = 0; i < rows.length; i++) {
        var r = rows[i];
        var tw = formatSize(r.finalW), th = formatSize(r.finalH);
        // PRIME sheets already include bleed -- show the sheet size as-is
        var rb = r.isPrime ? 0 : bleed;
        var bw = formatSize(finishedDim(r.finalW, rb)), bh = formatSize(finishedDim(r.finalH, rb));
        rawTw.push(tw); rawTh.push(th); rawBw.push(bw); rawBh.push(bh);
        if (tw.length > wTw) wTw = tw.length;
        if (th.length > wTh) wTh = th.length;
        if (bw.length > wBw) wBw = bw.length;
        if (bh.length > wBh) wBh = bh.length;
    }

    var out = [];
    var lastDocName = null;
    for (var j = 0; j < rows.length; j++) {
        var r2 = rows[j];
        if (lastDocName !== null && r2.docName !== lastDocName) {
            out.push({ sep: true });
        }

        var est = r2._estimate;
        var roll, run;
        if (est && est.ok) {
            roll = formatSize(est.rollSize) + "\"";
            run  = formatSize(est.runWithWaste) + "\"";
        } else {
            roll = "--";
            run  = "--";
        }

        out.push({
            sep: false,
            num:  r2.index + ".",
            name: r2.name,
            trim:  r2.isPrime ? "nested" : padLeft(rawTw[j], wTw) + "\" x " + padLeft(rawTh[j], wTh) + "\"",
            bleed: padLeft(rawBw[j], wBw) + "\" x " + padLeft(rawBh[j], wBh) + "\"",
            roll: roll,
            run:  run
        });
        lastDocName = r2.docName;
    }
    return out;
}

// Find the longest common prefix shared by all strings, trimmed back to a
// natural separator (_  -  space  .) so we don't cut mid-word.
function findCommonPrefix(strings) {
    if (strings.length < 2) return "";

    var prefix = strings[0];
    for (var i = 1; i < strings.length; i++) {
        while (strings[i].indexOf(prefix) !== 0) {
            prefix = prefix.substring(0, prefix.length - 1);
            if (prefix === "") return "";
        }
    }

    // If prefix doesn't end at a natural separator, trim back to the last one
    var separators = "_- .";
    var lastChar = prefix.charAt(prefix.length - 1);
    if (separators.indexOf(lastChar) === -1) {
        var lastSepIdx = -1;
        for (var idx = prefix.length - 1; idx >= 0; idx--) {
            if (separators.indexOf(prefix.charAt(idx)) !== -1) {
                lastSepIdx = idx;
                break;
            }
        }
        prefix = (lastSepIdx === -1) ? "" : prefix.substring(0, lastSepIdx + 1);
    }

    return prefix;
}

// ============================================================================
// CLIPBOARD HELPERS
// ============================================================================
function setClipboardText(text) {
    try {
        var tempTxt = new File(Folder.temp + "/_artboard_clip.txt");
        tempTxt.encoding = "UTF-8";
        tempTxt.open("w");
        tempTxt.write(text);
        tempTxt.close();

        var txtPath = tempTxt.fsName;

        if (File.fs === "Windows") {
            var tempVbs = new File(Folder.temp + "/_artboard_clip.vbs");
            tempVbs.encoding = "UTF-8";
            tempVbs.open("w");
            tempVbs.writeln('CreateObject("WScript.Shell").Run "cmd /c clip < " & Chr(34) & "' + txtPath + '" & Chr(34), 0, True');
            tempVbs.close();
            tempVbs.execute();
            $.sleep(800);
            tempVbs.remove();
        } else {
            var tempSh = new File(Folder.temp + "/_artboard_clip.command");
            tempSh.encoding = "UTF-8";
            tempSh.open("w");
            tempSh.writeln('#!/bin/bash');
            tempSh.writeln('pbcopy < "' + txtPath + '"');
            tempSh.close();
            tempSh.execute();
            $.sleep(800);
            tempSh.remove();
        }

        tempTxt.remove();
        return true;
    } catch (e) {
        return false;
    }
}

function showCopyFallbackDialog(text) {
    var dlg = new Window("dialog", "Copy this text");
    dlg.alignChildren = ["fill", "fill"];

    dlg.add("statictext", undefined, "Click in the box, press Ctrl+A then Ctrl+C, then paste into your spreadsheet:");

    var edit = dlg.add("edittext", undefined, text, {multiline: true, scrolling: true});
    edit.preferredSize = [500, 300];

    var btn = dlg.add("button", undefined, "Done", {name: "ok"});

    dlg.onShow = function() { edit.active = true; };
    dlg.show();
}

// ============================================================================
// COPY SHAPES TO CLIPBOARD
// ============================================================================
// Single-doc mode: shapes match each artboard's exact size and position.
// Multi-doc mode: shapes are arranged in a horizontal row in the active doc
// (since artboards from different docs can't share coordinate systems).
function copyShapesToClipboard(rows, multiDoc) {
    var doc = app.activeDocument;

    // Save current selection so we can restore after
    var prevSelection = [];
    for (var s = 0; s < doc.selection.length; s++) prevSelection.push(doc.selection[s]);

    // Stroke color appropriate to active doc's color mode
    var strokeColor;
    if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
        strokeColor = new CMYKColor();
        strokeColor.cyan = 0; strokeColor.magenta = 0; strokeColor.yellow = 0; strokeColor.black = 100;
    } else {
        strokeColor = new RGBColor();
        strokeColor.red = 0; strokeColor.green = 0; strokeColor.blue = 0;
    }

    var tempLayer = doc.layers.add();
    tempLayer.name = "_TEMP_ARTBOARD_SHAPES_";

    var createdRects = [];
    var padding = 36; // 0.5" between shapes when arranging in a row
    var currentX = 0;

    for (var i = 0; i < rows.length; i++) {
        var r = rows[i];
        var left, top, width, height;

        if (multiDoc) {
            // Arrange in a single horizontal row near origin
            width = r.origWpts;
            height = r.origHpts;
            left = currentX;
            top = height; // y of upper-left corner
            currentX += width + padding;
        } else {
            // Use original artboard position from the source doc
            var rect = r.origRect;
            left = rect[0];
            top = rect[1];
            width = r.origWpts;
            height = r.origHpts;
        }

        var newRect = tempLayer.pathItems.rectangle(top, left, width, height);
        newRect.name = r.artboardName || r.name;
        newRect.filled = false;
        newRect.stroked = true;
        newRect.strokeWidth = 1;
        newRect.strokeColor = strokeColor;
        createdRects.push(newRect);
    }

    // Select only the new rectangles
    doc.selection = null;
    for (var j = 0; j < createdRects.length; j++) createdRects[j].selected = true;
    app.redraw();

    app.executeMenuCommand("copy");

    // Cleanup
    tempLayer.remove();
    doc.selection = null;
    for (var k = 0; k < prevSelection.length; k++) {
        try { prevSelection[k].selected = true; } catch (e) {}
    }
}

// ============================================================================
// MATERIAL ESTIMATION
// ============================================================================
// (ESTIMATION_DEFAULTS is declared at top of file so it's available at startup)

// Compute the printed-material requirement for one panel.
//
// isPrime = true means the artboard is a nested PRIME sheet: bleed is already
// baked into the artboard, so no bleed is added and the artboard area is NOT
// installed coverage (it includes bleed plus the gaps between nested panels).
//
// All area results are in sq ft. The printed area is split into parts so the
// waste can be explained:
//   printed = installed + bleed + unused roll width + R.L. waste   (panels)
//   printed = sheet     + unused roll width + R.L. waste           (PRIME)
function calculatePanelEstimate(width, height, params, isPrime) {
    var bleed = isPrime ? 0 : params.bleed;
    var w1 = width + bleed * 2;
    var h1 = height + bleed * 2;
    var rollRequired = Math.min(w1, h1); // panel goes across the roll on its short edge
    var runLength = Math.max(w1, h1);    // long edge runs down the roll

    // Pick smallest roll where usable width fits the panel
    var rollSize = null;
    for (var i = 0; i < params.rollWidths.length; i++) {
        if (params.rollWidths[i] - params.pinchMargin >= rollRequired) {
            rollSize = params.rollWidths[i];
            break;
        }
    }

    var ok = (rollSize !== null);

    // If the chosen roll is also wide enough for the long edge, rotate the
    // panel so the long edge goes across and the short edge becomes the run.
    // (Happens when smaller rolls are unchecked, e.g. 48x38.5 on a 54" roll.)
    if (ok && rollSize - params.pinchMargin >= runLength) {
        var tmp = rollRequired;
        rollRequired = runLength;
        runLength = tmp;
    }

    var runWithWaste = runLength + params.rlWaste;
    var sheetSqFt = (w1 * h1) / 144;
    var installedSqFt = isPrime ? 0 : (width * height) / 144;
    var bleedSqFt = isPrime ? 0 : sheetSqFt - installedSqFt;
    var printedSqFt = ok ? (rollSize * runWithWaste) / 144 : 0;
    var sideWasteSqFt = ok ? (rollSize * runLength) / 144 - sheetSqFt : 0;
    var rlWasteSqFt = ok ? (rollSize * params.rlWaste) / 144 : 0;

    return {
        isPrime: !!isPrime,
        installedSqFt: installedSqFt,
        sheetSqFt: sheetSqFt,
        bleedSqFt: bleedSqFt,
        sideWasteSqFt: sideWasteSqFt,
        rlWasteSqFt: rlWasteSqFt,
        printedSqFt: printedSqFt,
        rollRequired: rollRequired,
        runLength: runLength,
        runWithWaste: runWithWaste,
        rollSize: rollSize,
        ok: ok
    };
}

// Aggregate one group of rows (all panel rows, or all PRIME rows).
function aggregateGroup(rows, params, isPrime) {
    var t = {
        count: 0, installed: 0, printed: 0, sheet: 0,
        bleed: 0, sideWaste: 0, rlWaste: 0, errorPanels: []
    };
    var rollUsage = {};

    for (var i = 0; i < rows.length; i++) {
        var r = rows[i];
        var est = calculatePanelEstimate(r.finalW, r.finalH, params, isPrime);
        r._estimate = est;

        t.count += 1;
        t.installed += est.installedSqFt;

        if (est.ok) {
            t.printed += est.printedSqFt;
            t.sheet += est.sheetSqFt;
            t.bleed += est.bleedSqFt;
            t.sideWaste += est.sideWasteSqFt;
            t.rlWaste += est.rlWasteSqFt;

            var key = est.rollSize.toString();
            if (!rollUsage[key]) rollUsage[key] = { rollSize: est.rollSize, count: 0, sqft: 0, runInches: 0 };
            rollUsage[key].count += 1;
            rollUsage[key].sqft += est.printedSqFt;
            rollUsage[key].runInches += est.runWithWaste;
        } else {
            t.errorPanels.push(r.name);
        }
    }

    var rollList = [];
    for (var k in rollUsage) {
        if (rollUsage.hasOwnProperty(k)) rollList.push(rollUsage[k]);
    }
    rollList.sort(function(a, b) { return a.rollSize - b.rollSize; });

    return { isPrime: isPrime, totals: t, rollList: rollList };
}

// Aggregate per-doc and overall estimate from the rows array.
// Panel files and PRIME (nested) files are kept in separate groups so the
// worst-case (no nesting) and nested figures can be compared side by side.
// When both are selected, the panel files' installed coverage is used as the
// installed figure for the PRIME too (same job, same panels).
function calculateOverallEstimate(rows, params) {
    var panelRows = [], primeRows = [];
    for (var i = 0; i < rows.length; i++) {
        if (rows[i].isPrime) primeRows.push(rows[i]);
        else panelRows.push(rows[i]);
    }

    var panel = panelRows.length ? aggregateGroup(panelRows, params, false) : null;
    var prime = primeRows.length ? aggregateGroup(primeRows, params, true) : null;

    var perDoc = {};
    var docOrder = [];
    for (var j = 0; j < rows.length; j++) {
        var r = rows[j];
        if (!perDoc[r.docName]) {
            perDoc[r.docName] = {
                docName: r.docName,
                docBaseName: r.docBaseName,
                isPrime: !!r.isPrime,
                panels: 0,
                installed: 0,
                printed: 0
            };
            docOrder.push(r.docName);
        }
        var b = perDoc[r.docName];
        b.panels += 1;
        b.installed += r._estimate.installedSqFt;
        b.printed += r._estimate.printedSqFt;
    }

    return {
        panel: panel,
        prime: prime,
        perDoc: perDoc,
        docOrder: docOrder,
        params: params
    };
}

// Build the coverage report as a list of sections. Each section is
// { title, rows: [ { label, value, bold, indent } ], note }.
// The dialog text, plain-text email, and HTML email all render from this so
// the three always show the same numbers.
function buildCoverageSections(estimate) {
    var p = estimate.params;
    var sections = [];
    var panel = estimate.panel;
    var prime = estimate.prime;

    function row(label, value, bold, indent) {
        return { label: label, value: value, bold: !!bold, indent: !!indent };
    }
    function sq(n) { return formatSqFt(n) + " sq ft"; }
    function pctOf(part, whole) { return whole > 0 ? formatPct(part / whole) : "--"; }

    function rollRows(group) {
        var out = [];
        for (var j = 0; j < group.rollList.length; j++) {
            var rl = group.rollList[j];
            out.push(row(rl.rollSize + "\" roll", inchesToFeet(rl.runInches) + " ft  (" + sq(rl.sqft) + ")", false, true));
        }
        return out;
    }

    if (panel) {
        var t = panel.totals;
        var waste = t.printed - t.installed;
        var rows = [];
        rows.push(row("Panels", t.count.toString()));
        rows.push(row("Installed", sq(t.installed), true));
        rows.push(row("Printed", sq(t.printed), true));
        rows.push(row("Utilization", pctOf(t.installed, t.printed)));
        rows.push(row("Waste", sq(waste) + "  (" + pctOf(waste, t.printed) + ")"));
        rows.push(row("Bleed (" + p.bleed + "\"/side)", sq(t.bleed) + "  (" + pctOf(t.bleed, t.printed) + ")", false, true));
        rows.push(row("Unused roll width", sq(t.sideWaste) + "  (" + pctOf(t.sideWaste, t.printed) + ")", false, true));
        rows.push(row("R.L. waste (" + p.rlWaste + "\" x " + t.count + ")", sq(t.rlWaste) + "  (" + pctOf(t.rlWaste, t.printed) + ")", false, true));
        rows.push(row("Roll length needed", ""));
        rows = rows.concat(rollRows(panel));
        sections.push({
            title: prime ? "Panels - no nesting (worst case)" : "Total Coverage - panels, no nesting",
            rows: rows
        });
    }

    if (prime) {
        var tp = prime.totals;
        var installed = panel ? panel.totals.installed : 0;
        var prow = [];
        prow.push(row("Nested sheets", tp.count.toString()));
        if (panel) prow.push(row("Installed (from panel files)", sq(installed), true));
        prow.push(row("Printed", sq(tp.printed), true));
        if (panel) {
            var pw = tp.printed - installed;
            var inNest = tp.sheet - installed;
            prow.push(row("Utilization", pctOf(installed, tp.printed)));
            prow.push(row("Waste", sq(pw) + "  (" + pctOf(pw, tp.printed) + ")"));
            prow.push(row("Bleed + gaps inside nests", sq(inNest) + "  (" + pctOf(inNest, tp.printed) + ")", false, true));
        } else {
            prow.push(row("Sheet area (incl. bleed)", sq(tp.sheet)));
        }
        prow.push(row("Unused roll width", sq(tp.sideWaste) + "  (" + pctOf(tp.sideWaste, tp.printed) + ")", false, true));
        prow.push(row("R.L. waste (" + p.rlWaste + "\" x " + tp.count + ")", sq(tp.rlWaste) + "  (" + pctOf(tp.rlWaste, tp.printed) + ")", false, true));
        prow.push(row("Roll length needed", ""));
        prow = prow.concat(rollRows(prime));
        sections.push({
            title: "PRIME - nested (expected)",
            rows: prow,
            note: panel ? "" : "Bleed is already in the PRIME artboards, so none is added. Select the panel file(s) too to get installed coverage and utilization."
        });
    }

    if (panel && prime) {
        var saved = panel.totals.printed - prime.totals.printed;
        sections.push({
            title: "Nesting Savings",
            rows: [
                row("Printed, no nesting", sq(panel.totals.printed)),
                row("Printed, nested", sq(prime.totals.printed)),
                row("Saved by nesting", sq(saved) + "  (" + pctOf(saved, panel.totals.printed) + ")", true)
            ]
        });
    }

    var errs = [];
    if (panel) errs = errs.concat(panel.totals.errorPanels);
    if (prime) errs = errs.concat(prime.totals.errorPanels);
    if (errs.length > 0) {
        var erows = [];
        for (var e = 0; e < errs.length; e++) erows.push(row("- " + errs[e], ""));
        sections.push({ title: "WARNING - too wide for any available roll", rows: erows, warning: true });
    }

    return sections;
}

// Render sections as aligned plain text.
function renderSectionsText(sections, underline) {
    var lines = [];
    for (var s = 0; s < sections.length; s++) {
        var sec = sections[s];
        var w = 0;
        for (var i = 0; i < sec.rows.length; i++) {
            var len = sec.rows[i].label.length + (sec.rows[i].indent ? 2 : 0);
            if (sec.rows[i].value !== "" && len > w) w = len;
        }
        if (s > 0) lines.push("");
        lines.push(sec.title.toUpperCase());
        if (underline) {
            var u = "";
            while (u.length < sec.title.length) u += "-";
            lines.push(u);
        }
        for (var j = 0; j < sec.rows.length; j++) {
            var r = sec.rows[j];
            var label = (r.indent ? "  " : "") + r.label;
            if (r.value === "") lines.push(label + (r.label.charAt(0) === "-" ? "" : ":"));
            else lines.push(padRight(label + ":", w + 2) + r.value);
        }
        if (sec.note) lines.push("(" + sec.note + ")");
    }
    return lines.join("\n");
}

// Format a sq ft number (2 decimals)
function formatSqFt(n) {
    return (Math.round(n * 100) / 100).toFixed(2);
}

// Format a percentage (1 decimal)
function formatPct(n) {
    return (Math.round(n * 1000) / 10).toFixed(1) + "%";
}

// Inches to feet, 1 decimal
function inchesToFeet(n) {
    return (Math.round((n / 12) * 10) / 10).toFixed(1);
}

// Build the in-dialog estimate summary text.
function buildEstimateSummary(estimate, multiDoc) {
    var head = "";
    if (multiDoc) {
        var lines = [];
        for (var i = 0; i < estimate.docOrder.length; i++) {
            var d = estimate.perDoc[estimate.docOrder[i]];
            if (d.isPrime) {
                lines.push(d.docBaseName + ":  " + d.panels + " nested sheets   Printed: " + formatSqFt(d.printed) + " sq ft   [PRIME]");
            } else {
                lines.push(d.docBaseName + ":  " + d.panels + " panels   Installed: " + formatSqFt(d.installed) + " sq ft   Printed: " + formatSqFt(d.printed) + " sq ft");
            }
        }
        head = lines.join("\n") + "\n\n";
    }
    return head + renderSectionsText(buildCoverageSections(estimate), false);
}

// Compute the project name and per-section labels from doc base names
function getProjectInfo(estimate, multiDoc) {
    var info = { projectName: "", sectionNames: {} };
    if (multiDoc) {
        var baseNames = [];
        for (var i = 0; i < estimate.docOrder.length; i++) {
            baseNames.push(estimate.perDoc[estimate.docOrder[i]].docBaseName);
        }
        info.projectName = findCommonPrefix(baseNames).replace(/[_\-\s\.]+$/, "");
        if (!info.projectName) info.projectName = "Multiple files";
        for (var d = 0; d < estimate.docOrder.length; d++) {
            var bn = estimate.perDoc[estimate.docOrder[d]].docBaseName;
            var section = bn;
            if (info.projectName && bn.indexOf(info.projectName) === 0) {
                section = bn.substring(info.projectName.length).replace(/^[_\-\s\.]+/, "");
            }
            if (!section) section = bn;
            info.sectionNames[estimate.docOrder[d]] = section;
        }
    } else {
        info.projectName = estimate.perDoc[estimate.docOrder[0]].docBaseName;
    }
    return info;
}

// Build email-friendly plain text -- coverage + roll info, and optionally the
// per-panel size table (trim / w-bleed / roll / run).
function buildEmailText(rows, estimate, multiDoc, includeSizes, bleed) {
    var info = getProjectInfo(estimate, multiDoc);
    var lines = [];

    lines.push("Project: " + info.projectName);
    lines.push("");

    if (includeSizes) {
        var drows = buildDisplayRows(rows, bleed);

        // Column widths for aligned plain text
        var wNum = 2, wName = 5, wTrim = 4, wBleed = 8, wRoll = 4;
        for (var di = 0; di < drows.length; di++) {
            var d0 = drows[di];
            if (d0.sep) continue;
            if (d0.num.length   > wNum)   wNum   = d0.num.length;
            if (d0.name.length  > wName)  wName  = d0.name.length;
            if (d0.trim.length  > wTrim)  wTrim  = d0.trim.length;
            if (d0.bleed.length > wBleed) wBleed = d0.bleed.length;
            if (d0.roll.length  > wRoll)  wRoll  = d0.roll.length;
        }

        lines.push("PANEL SIZES (x10, rounded up to 0.5\")");
        lines.push("--------------------------------------");
        lines.push(padRight("#", wNum) + "  " + padRight("Panel", wName) + "  " +
                   padRight("Trim", wTrim) + "  " + padRight("W/ Bleed", wBleed) + "  " +
                   padRight("Roll", wRoll) + "  " + "Run");
        for (var dj = 0; dj < drows.length; dj++) {
            var d1 = drows[dj];
            if (d1.sep) {
                lines.push("");
                continue;
            }
            lines.push(padRight(d1.num, wNum) + "  " + padRight(d1.name, wName) + "  " +
                       padRight(d1.trim, wTrim) + "  " + padRight(d1.bleed, wBleed) + "  " +
                       padRight(d1.roll, wRoll) + "  " + d1.run);
        }
        lines.push("");
    }

    lines.push(renderSectionsText(buildCoverageSections(estimate), true));

    return lines.join("\n");
}

// Escape special HTML characters
function escapeHtml(s) {
    if (s === undefined || s === null) return "";
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// Build rich-text HTML -- coverage + roll info, and optionally the per-panel
// size table. Pasted into Outlook/Gmail/Apple Mail, this renders as a clean
// styled block.
function buildEmailHtml(rows, estimate, multiDoc, includeSizes, bleed) {
    var info = getProjectInfo(estimate, multiDoc);

    var h3 = 'style="margin:20px 0 6px;font-size:13px;color:#444;text-transform:uppercase;letter-spacing:0.6px;font-weight:bold;border-bottom:1px solid #ddd;padding-bottom:4px;"';
    var tableStyle = 'style="border-collapse:collapse;font-size:13px;margin:0 0 8px 0;"';
    var thStyle = 'style="text-align:left;padding:6px 14px 6px 0;background:#f4f4f4;border-bottom:1px solid #ccc;font-weight:bold;"';
    var thRight = 'style="text-align:right;padding:6px 0 6px 14px;background:#f4f4f4;border-bottom:1px solid #ccc;font-weight:bold;"';
    var tdStyle = 'style="padding:5px 14px 5px 0;border-bottom:1px solid #eee;"';
    var tdRight = 'style="padding:5px 0 5px 14px;border-bottom:1px solid #eee;text-align:right;"';

    var html = '';
    html += '<div style="font-family:Arial,Helvetica,sans-serif;color:#1a1a1a;font-size:13px;line-height:1.5;">';

    html += '<div style="margin:0 0 12px;color:#222;font-size:14px;"><b>Project:</b> ' + escapeHtml(info.projectName) + '</div>';

    if (includeSizes) {
        var drows = buildDisplayRows(rows, bleed);
        html += '<div ' + h3 + '>Panel Sizes <span style="font-weight:normal;text-transform:none;letter-spacing:0;">(x10, rounded up to 0.5&quot;)</span></div>';
        html += '<table ' + tableStyle + '>';
        html += '<tr>';
        html += '<th ' + thStyle + '>#</th>';
        html += '<th ' + thStyle + '>Panel</th>';
        html += '<th ' + thRight + '>Trim</th>';
        html += '<th ' + thRight + '>W/ Bleed</th>';
        html += '<th ' + thRight + '>Roll</th>';
        html += '<th ' + thRight + '>Run</th>';
        html += '</tr>';
        for (var di = 0; di < drows.length; di++) {
            var d = drows[di];
            if (d.sep) {
                // Spacer row between document sets
                html += '<tr><td colspan="6" style="padding:8px 0;border-bottom:none;"></td></tr>';
                continue;
            }
            html += '<tr>';
            html += '<td ' + tdStyle + '>' + escapeHtml(d.num) + '</td>';
            html += '<td ' + tdStyle + '>' + escapeHtml(d.name) + '</td>';
            html += '<td ' + tdRight + ' nowrap>' + escapeHtml(d.trim) + '</td>';
            html += '<td ' + tdRight + ' nowrap>' + escapeHtml(d.bleed) + '</td>';
            html += '<td ' + tdRight + '>' + escapeHtml(d.roll) + '</td>';
            html += '<td ' + tdRight + '>' + escapeHtml(d.run) + '</td>';
            html += '</tr>';
        }
        html += '</table>';
    }

    var h3warn = 'style="margin:20px 0 6px;font-size:13px;color:#c00;text-transform:uppercase;letter-spacing:0.6px;font-weight:bold;"';
    var tdIndent = 'style="padding:4px 14px 4px 18px;border-bottom:1px solid #eee;color:#555;"';
    var tdSub = 'style="padding:8px 14px 3px 0;color:#666;font-style:italic;"';
    var sections = buildCoverageSections(estimate);
    for (var s = 0; s < sections.length; s++) {
        var sec = sections[s];
        html += '<div ' + (sec.warning ? h3warn : h3) + '>' + escapeHtml(sec.title) + '</div>';
        html += '<table ' + tableStyle + '>';
        for (var ri = 0; ri < sec.rows.length; ri++) {
            var r = sec.rows[ri];
            if (r.value === "") {
                html += '<tr><td colspan="2" ' + (sec.warning ? 'style="color:#c00;padding:2px 0;"' : tdSub) + '>' + escapeHtml(r.label) + '</td></tr>';
                continue;
            }
            var val = escapeHtml(r.value);
            if (r.bold) val = '<b>' + val + '</b>';
            html += '<tr><td ' + (r.indent ? tdIndent : tdStyle) + '>' + escapeHtml(r.label) + '</td><td ' + tdRight + ' nowrap>' + val + '</td></tr>';
        }
        html += '</table>';
        if (sec.note) html += '<div style="color:#666;font-size:12px;margin:0 0 8px;">' + escapeHtml(sec.note) + '</div>';
    }

    html += '</div>';
    return html;
}

// Set both HTML and plain-text formats on the Windows clipboard via PowerShell
// so emails paste as rich text but plain-text editors still get readable text.
function setClipboardHtml(html, plainText) {
    if (File.fs !== "Windows") {
        // Mac / other: fall back to plain text
        return setClipboardText(plainText);
    }

    try {
        var tempHtml = new File(Folder.temp + "/_artboard_clip.html");
        tempHtml.encoding = "UTF-8";
        tempHtml.open("w");
        tempHtml.write(html);
        tempHtml.close();

        var tempText = new File(Folder.temp + "/_artboard_clip_text.txt");
        tempText.encoding = "UTF-8";
        tempText.open("w");
        tempText.write(plainText);
        tempText.close();

        // PowerShell script that loads both files and puts them on the clipboard
        // as a single DataObject containing HTML + plain text formats.
        //
        // IMPORTANT: the Windows "HTML Format" clipboard type requires a CF_HTML
        // header with byte offsets (StartHTML/EndHTML/StartFragment/EndFragment).
        // Without it, browsers (Gmail/Chrome) reject the HTML and paste falls
        // back to plain text. Offsets are UTF-8 BYTE positions, so the payload
        // is written as a UTF-8 byte stream, not a .NET string.
        var ps = "";
        ps += "Add-Type -AssemblyName System.Windows.Forms\n";
        ps += "$html = Get-Content -Raw -Encoding UTF8 -Path '" + tempHtml.fsName.replace(/'/g, "''") + "'\n";
        ps += "$text = Get-Content -Raw -Encoding UTF8 -Path '" + tempText.fsName.replace(/'/g, "''") + "'\n";
        ps += "$enc = [System.Text.Encoding]::UTF8\n";
        ps += "$pre = '<html><body><!--StartFragment-->'\n";
        ps += "$post = '<!--EndFragment--></body></html>'\n";
        ps += "$tpl = \"Version:0.9`r`nStartHTML:{0:D10}`r`nEndHTML:{1:D10}`r`nStartFragment:{2:D10}`r`nEndFragment:{3:D10}`r`n\"\n";
        ps += "$hdrLen = $enc.GetByteCount(($tpl -f 0,0,0,0))\n";
        ps += "$startFrag = $hdrLen + $enc.GetByteCount($pre)\n";
        ps += "$endFrag = $startFrag + $enc.GetByteCount($html)\n";
        ps += "$endHtml = $endFrag + $enc.GetByteCount($post)\n";
        ps += "$cf = ($tpl -f $hdrLen, $endHtml, $startFrag, $endFrag) + $pre + $html + $post\n";
        ps += "$ms = New-Object System.IO.MemoryStream (,$enc.GetBytes($cf))\n";
        ps += "$do = New-Object System.Windows.Forms.DataObject\n";
        ps += "$do.SetText($text)\n";
        ps += "$do.SetData('HTML Format', $ms)\n";
        ps += "[System.Windows.Forms.Clipboard]::SetDataObject($do, $true)\n";

        var tempPs = new File(Folder.temp + "/_artboard_clip.ps1");
        tempPs.encoding = "UTF-8";
        tempPs.open("w");
        tempPs.write(ps);
        tempPs.close();

        // VBS launches PowerShell hidden and waits for it
        var tempVbs = new File(Folder.temp + "/_artboard_clip_html.vbs");
        tempVbs.encoding = "UTF-8";
        tempVbs.open("w");
        tempVbs.writeln('CreateObject("WScript.Shell").Run "powershell.exe -sta -NoProfile -ExecutionPolicy Bypass -File " & Chr(34) & "' + tempPs.fsName + '" & Chr(34), 0, True');
        tempVbs.close();
        tempVbs.execute();

        // PowerShell + Add-Type takes time to spin up
        $.sleep(2500);

        tempVbs.remove();
        tempPs.remove();
        tempHtml.remove();
        tempText.remove();
        return true;
    } catch (e) {
        // Fall back to plain text
        return setClipboardText(plainText);
    }
}

// ============================================================================
// MIRROR DRIVER <-> PASSENGER
// ============================================================================
// Match the case pattern of `template` onto `target` (e.g. "DRIVER" -> "PSSNGR",
// "Driver" -> "Pssngr", "driver" -> "pssngr").
function matchCase(template, target) {
    if (template === template.toUpperCase()) return target.toUpperCase();
    if (template === template.toLowerCase()) return target.toLowerCase();
    if (template.charAt(0) === template.charAt(0).toUpperCase()) {
        return target.charAt(0).toUpperCase() + target.substring(1).toLowerCase();
    }
    return target.toLowerCase();
}

// Inspect the doc name and figure out which side it is + likely target spellings
function getActiveSideInfo(doc) {
    var name = doc.name;
    var m;
    if ((m = name.match(/Driver/i))) {
        return { sourceTerm: m[0], targetTerms: ['Pssngr', 'Passenger'] };
    }
    if ((m = name.match(/Pssngr/i))) {
        return { sourceTerm: m[0], targetTerms: ['Driver'] };
    }
    if ((m = name.match(/Passenger/i))) {
        return { sourceTerm: m[0], targetTerms: ['Driver'] };
    }
    return null;
}

// Find the matching opposite-side file by substituting the side term in the
// source doc name and looking for a matching open document.
function findMirrorTargetDoc(srcDoc, sideInfo) {
    var srcName = srcDoc.name;
    for (var i = 0; i < sideInfo.targetTerms.length; i++) {
        var targetTerm = matchCase(sideInfo.sourceTerm, sideInfo.targetTerms[i]);
        var candidateName = srcName.replace(new RegExp(sideInfo.sourceTerm, 'i'), targetTerm);
        for (var j = 0; j < app.documents.length; j++) {
            if (app.documents[j].name === candidateName && app.documents[j] !== srcDoc) {
                return { doc: app.documents[j], targetTerm: targetTerm };
            }
        }
    }
    return null;
}

// Replace every case-insensitive occurrence of sourceTerm in name with targetTerm,
// preserving the matched substring's case style on each replacement.
function swapSideInName(name, sourceTerm, targetTerm) {
    var pattern = new RegExp(sourceTerm, 'gi');
    var result = '';
    var lastIndex = 0;
    var match;
    while ((match = pattern.exec(name)) !== null) {
        result += name.substring(lastIndex, match.index);
        result += matchCase(match[0], targetTerm);
        lastIndex = match.index + match[0].length;
        // Avoid infinite loop on zero-length matches
        if (match[0].length === 0) pattern.lastIndex++;
    }
    result += name.substring(lastIndex);
    return result;
}

// Mirror all artboards from the active doc to the matching opposite-side doc.
// - Mirror axis: horizontal center of source artboards' bounding box.
// - Skips artboards literally named "proof".
// - Adds to existing artboards in target (does not delete).
// - Updates artboard names by swapping the side keyword (case-preserving).
function performMirrorDP() {
    var srcDoc = app.activeDocument;
    var sideInfo = getActiveSideInfo(srcDoc);
    if (!sideInfo) {
        alert("Active document name doesn't include 'Driver', 'Pssngr', or 'Passenger'.\nMirror D/P needs the file name to identify which side to mirror.");
        return;
    }

    var targetInfo = findMirrorTargetDoc(srcDoc, sideInfo);
    if (!targetInfo) {
        var expected = srcDoc.name.replace(new RegExp(sideInfo.sourceTerm, 'i'), matchCase(sideInfo.sourceTerm, sideInfo.targetTerms[0]));
        alert("Couldn't find the opposite-side file.\nExpected an open document named:\n  " + expected + "\n\nMake sure both files are open.");
        return;
    }

    var targetDoc = targetInfo.doc;
    var targetTerm = targetInfo.targetTerm;

    // Activate source so artboards collection returns its data
    app.activeDocument = srcDoc;
    app.redraw();
    $.sleep(100);

    // Snapshot source artboards
    var srcArtboards = [];
    var minLeft = Infinity, maxRight = -Infinity;
    for (var k = 0; k < srcDoc.artboards.length; k++) {
        var ab = srcDoc.artboards[k];
        if (ab.name.toLowerCase() === 'proof') continue;
        var rect = ab.artboardRect;
        srcArtboards.push({ name: ab.name, rect: [rect[0], rect[1], rect[2], rect[3]] });
        if (rect[0] < minLeft) minLeft = rect[0];
        if (rect[2] > maxRight) maxRight = rect[2];
    }

    if (srcArtboards.length === 0) {
        alert("No mirrorable artboards in source (all are named 'proof' or there are none).");
        return;
    }

    var mirrorX = (minLeft + maxRight) / 2;

    // Confirm
    var msg = "Mirror " + srcArtboards.length + " artboard(s)\n";
    msg += "  from:  " + srcDoc.name + "\n";
    msg += "  to:    " + targetDoc.name + "\n\n";
    msg += "Existing artboards in the target document will be preserved.\nContinue?";
    if (!confirm(msg)) return;

    // Switch to target and add mirrored artboards
    app.activeDocument = targetDoc;
    app.redraw();
    $.sleep(100);

    var added = 0;
    var errors = [];
    for (var m = 0; m < srcArtboards.length; m++) {
        var src = srcArtboards[m];
        var newL = 2 * mirrorX - src.rect[2];
        var newR = 2 * mirrorX - src.rect[0];
        var newT = src.rect[1];
        var newB = src.rect[3];
        var newName = swapSideInName(src.name, sideInfo.sourceTerm, targetTerm);

        try {
            var newAb = targetDoc.artboards.add([newL, newT, newR, newB]);
            newAb.name = newName;
            added++;
        } catch (e) {
            errors.push(src.name + ": " + e.toString());
        }
    }

    app.redraw();

    var doneMsg = "Mirrored " + added + " artboard(s) into:\n  " + targetDoc.name;
    if (errors.length > 0) {
        doneMsg += "\n\nErrors on " + errors.length + " artboard(s):\n" + errors.join("\n");
    }
    alert(doneMsg);
}

// ============================================================================
// RESULTS DIALOG
// ============================================================================
function showResults(rows, multiDoc) {
    var dlg = new Window("dialog", "Artboard Sizes & Estimate (" + rows.length + ")");
    dlg.alignChildren = ["fill", "fill"];
    dlg.spacing = 10;

    // ---- Estimation parameters panel ----
    var paramsPanel = dlg.add("panel", undefined, "Estimation Parameters");
    paramsPanel.orientation = "column";
    paramsPanel.alignChildren = "left";
    paramsPanel.margins = 12;
    paramsPanel.spacing = 8;

    var paramsRow1 = paramsPanel.add("group");
    paramsRow1.orientation = "row";
    paramsRow1.spacing = 12;

    paramsRow1.add("statictext", undefined, "Bleed (in/side):");
    var bleedField = paramsRow1.add("edittext", undefined, ESTIMATION_DEFAULTS.bleed.toString());
    bleedField.preferredSize.width = 50;

    paramsRow1.add("statictext", undefined, "R.L. Waste (in):");
    var wasteField = paramsRow1.add("edittext", undefined, ESTIMATION_DEFAULTS.rlWaste.toString());
    wasteField.preferredSize.width = 50;

    paramsRow1.add("statictext", undefined, "Pinch Margin (in):");
    var marginField = paramsRow1.add("edittext", undefined, ESTIMATION_DEFAULTS.pinchMargin.toString());
    marginField.preferredSize.width = 50;

    var resetBtn = paramsRow1.add("button", undefined, "Reset");
    resetBtn.preferredSize.width = 60;

    // Roll-availability checkboxes -- uncheck a size to force panels onto the
    // next available larger roll
    var paramsRow2 = paramsPanel.add("group");
    paramsRow2.orientation = "row";
    paramsRow2.spacing = 12;
    paramsRow2.add("statictext", undefined, "Available rolls:");
    var rollCheckboxes = [];
    for (var rIdx = 0; rIdx < ESTIMATION_DEFAULTS.rollWidths.length; rIdx++) {
        var rw = ESTIMATION_DEFAULTS.rollWidths[rIdx];
        var rcb = paramsRow2.add("checkbox", undefined, rw + '"');
        rcb.value = true;
        rollCheckboxes.push({ width: rw, checkbox: rcb });
    }

    // Custom (unusual) roll size -- enabled when checked and the field has a
    // valid positive number. Gets merged into the sorted roll list so panels
    // pick the smallest fitting roll as usual.
    var customRollCheckbox = paramsRow2.add("checkbox", undefined, "Custom:");
    customRollCheckbox.value = false;
    var customRollField = paramsRow2.add("edittext", undefined, "");
    customRollField.preferredSize.width = 50;
    paramsRow2.add("statictext", undefined, '"');

    // ---- Artboard list ----
    var listLabel = dlg.add("statictext", undefined, "Artboard sizes (x10, rounded up to 0.5\")   [ trim -> w/ bleed ]   PRIME files: sheet size, bleed already included");

    // Real multi-column table: column boundaries are fixed in pixels, so the
    // columns stay aligned regardless of the platform font. Widths are sized to
    // the content so nothing clips.
    var initRows = buildDisplayRows(rows, ESTIMATION_DEFAULTS.bleed);
    var maxLen = { num: 1, name: 5, trim: 4, bleed: 8, roll: 4, run: 3 };
    for (var mi = 0; mi < initRows.length; mi++) {
        var dr = initRows[mi];
        if (dr.sep) continue;
        if (dr.num.length   > maxLen.num)   maxLen.num   = dr.num.length;
        if (dr.name.length  > maxLen.name)  maxLen.name  = dr.name.length;
        if (dr.trim.length  > maxLen.trim)  maxLen.trim  = dr.trim.length;
        if (dr.bleed.length > maxLen.bleed) maxLen.bleed = dr.bleed.length;
        if (dr.roll.length  > maxLen.roll)  maxLen.roll  = dr.roll.length;
        if (dr.run.length   > maxLen.run)   maxLen.run   = dr.run.length;
    }
    var CH = 8; // approx px per character
    var colWidths = [
        Math.max(34, maxLen.num  * CH + 10),
        Math.max(70, maxLen.name * CH + 16),
        Math.max(80, maxLen.trim * CH + 16),
        Math.max(80, maxLen.bleed* CH + 16),
        Math.max(50, maxLen.roll * CH + 12),
        Math.max(50, maxLen.run  * CH + 12)
    ];
    var listW = 24;
    for (var cwI = 0; cwI < colWidths.length; cwI++) listW += colWidths[cwI];

    var list = dlg.add("listbox", undefined, [], {
        numberOfColumns: 6,
        showHeaders: true,
        columnTitles: ["#", "Panel", "Trim", "W/ Bleed", "Roll", "Run"],
        columnWidths: colWidths
    });
    list.preferredSize = [Math.min(listW, 900), 300];

    // Fill (and refill) the table from the current bleed value.
    function fillList(bleed) {
        list.removeAll();
        var drows = buildDisplayRows(rows, bleed);
        for (var di = 0; di < drows.length; di++) {
            var d = drows[di];
            if (d.sep) {
                list.add("item", ""); // blank separator row
                continue;
            }
            var it = list.add("item", d.num);
            it.subItems[0].text = d.name;
            it.subItems[1].text = d.trim;
            it.subItems[2].text = d.bleed;
            it.subItems[3].text = d.roll;
            it.subItems[4].text = d.run;
        }
    }

    // ---- Estimate summary panel ----
    var summaryPanel = dlg.add("panel", undefined, "Material Estimate");
    summaryPanel.orientation = "column";
    summaryPanel.alignChildren = "fill";
    summaryPanel.margins = 12;

    var summaryText = summaryPanel.add("edittext", undefined, "", {multiline: true, scrolling: true, readonly: true});
    summaryText.preferredSize = [600, 280];

    // ---- Buttons ----
    var btnGroup = dlg.add("group");
    btnGroup.alignment = "right";

    var includeSizesCheckbox = btnGroup.add("checkbox", undefined, "Include panel sizes");
    includeSizesCheckbox.value = true;
    var copyEmailBtn = btnGroup.add("button", undefined, "Copy for Email");
    var copyShapesBtn = btnGroup.add("button", undefined, "Copy Shapes");
    var mirrorBtn = btnGroup.add("button", undefined, "Mirror D/P");
    var closeBtn = btnGroup.add("button", undefined, "Close", {name: "ok"});

    // ---- State + recalc ----
    var currentEstimate = null;

    function readParams() {
        var b = parseFloat(bleedField.text);
        var w = parseFloat(wasteField.text);
        var m = parseFloat(marginField.text);
        if (isNaN(b)) b = ESTIMATION_DEFAULTS.bleed;
        if (isNaN(w)) w = ESTIMATION_DEFAULTS.rlWaste;
        if (isNaN(m)) m = ESTIMATION_DEFAULTS.pinchMargin;

        // Only include checked roll sizes
        var rolls = [];
        for (var r = 0; r < rollCheckboxes.length; r++) {
            if (rollCheckboxes[r].checkbox.value) rolls.push(rollCheckboxes[r].width);
        }

        // Include the custom roll if the checkbox is on and the field is a
        // valid positive number
        if (customRollCheckbox.value) {
            var customW = parseFloat(customRollField.text);
            if (!isNaN(customW) && customW > 0) {
                rolls.push(customW);
            }
        }

        // Sort ascending so calculatePanelEstimate picks the smallest fitting roll
        rolls.sort(function(a, b) { return a - b; });

        return {
            bleed: b,
            rlWaste: w,
            pinchMargin: m,
            rollWidths: rolls
        };
    }

    function refresh() {
        var params = readParams();
        currentEstimate = calculateOverallEstimate(rows, params);
        summaryText.text = buildEstimateSummary(currentEstimate, multiDoc);
        // Rebuild the size table so the w/ bleed, Roll, and Run columns track
        // the current parameters
        fillList(params.bleed);
    }

    bleedField.onChange = refresh;
    wasteField.onChange = refresh;
    marginField.onChange = refresh;
    bleedField.onChanging = refresh;
    wasteField.onChanging = refresh;
    marginField.onChanging = refresh;

    for (var rcbIdx = 0; rcbIdx < rollCheckboxes.length; rcbIdx++) {
        rollCheckboxes[rcbIdx].checkbox.onClick = refresh;
    }

    // Auto-enable the custom checkbox when a valid number is typed, and
    // refresh on every change so the estimate updates live.
    customRollCheckbox.onClick = refresh;
    customRollField.onChange = function() {
        var v = parseFloat(customRollField.text);
        if (!isNaN(v) && v > 0) customRollCheckbox.value = true;
        refresh();
    };
    customRollField.onChanging = function() {
        var v = parseFloat(customRollField.text);
        if (!isNaN(v) && v > 0) customRollCheckbox.value = true;
        refresh();
    };

    resetBtn.onClick = function() {
        bleedField.text = ESTIMATION_DEFAULTS.bleed.toString();
        wasteField.text = ESTIMATION_DEFAULTS.rlWaste.toString();
        marginField.text = ESTIMATION_DEFAULTS.pinchMargin.toString();
        for (var i = 0; i < rollCheckboxes.length; i++) {
            rollCheckboxes[i].checkbox.value = true;
        }
        customRollCheckbox.value = false;
        customRollField.text = "";
        refresh();
    };

    refresh();

    // ---- Button handlers ----
    copyEmailBtn.onClick = function() {
        if (!currentEstimate) refresh();
        var includeSizes = includeSizesCheckbox.value;
        var bleed = readParams().bleed;
        var emailText = buildEmailText(rows, currentEstimate, multiDoc, includeSizes, bleed);
        var emailHtml = buildEmailHtml(rows, currentEstimate, multiDoc, includeSizes, bleed);

        // Rich HTML + plain-text on the clipboard: email clients paste the
        // styled tables, plain-text editors get the aligned text version.
        var ok = setClipboardHtml(emailHtml, emailText);
        if (ok) {
            alert("Email-ready estimate copied to clipboard." +
                  (includeSizes ? "\nIncludes the panel size table." : "") +
                  "\nPaste into your email.");
        } else {
            showCopyFallbackDialog(emailText);
        }
    };

    copyShapesBtn.onClick = function() {
        try {
            copyShapesToClipboard(rows, multiDoc);
            var note = multiDoc
                ? "Shapes were arranged in a row in the active document, then copied. Paste into the destination doc, rearrange, and convert to artboards."
                : "Shapes match the artboard sizes and positions exactly. Use Edit > Paste in Place in the destination doc to keep positions, then convert to artboards.";
            alert("Copied " + rows.length + " shape(s) to clipboard.\n\n" + note);
        } catch (e) {
            alert("Could not copy shapes:\n" + e);
        }
    };

    mirrorBtn.onClick = function() {
        try {
            performMirrorDP();
        } catch (e) {
            alert("Mirror D/P failed:\n" + e.toString());
        }
    };

    closeBtn.onClick = function() { dlg.close(); };

    dlg.show();
}
