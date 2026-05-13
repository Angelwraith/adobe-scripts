/*@METADATA{
  "name": "List Artboard Sizes",
  "description": "Lists all artboards across one or more open documents with sizes multiplied by 10 and rounded up to the nearest 0.5 inch. Includes copy buttons for spreadsheet TSV and shape duplicates.",
  "version": "1.2",
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
    rollWidths: [36, 54, 60]        // available stocked rolls
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
    // Filter out any documents with "proof" in the name — never useful for sizing
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
                finalH: heightRounded
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

function padRight(str, len) {
    var out = str;
    while (out.length < len) out += " ";
    return out;
}

function buildDisplayText(rows) {
    var maxNameLen = 0;
    for (var i = 0; i < rows.length; i++) {
        if (rows[i].name.length > maxNameLen) maxNameLen = rows[i].name.length;
    }

    var lines = [];
    lines.push("Artboard sizes (x10, rounded up to 0.5\"):");
    lines.push("");
    var lastDocName = null;
    for (var j = 0; j < rows.length; j++) {
        var r = rows[j];
        if (lastDocName !== null && r.docName !== lastDocName) {
            lines.push(""); // separator between document sets
        }
        var paddedName = padRight(r.name, maxNameLen);
        lines.push(r.index + ". " + paddedName + "   " + formatSize(r.finalW) + "\" x " + formatSize(r.finalH) + "\"");
        lastDocName = r.docName;
    }
    return lines.join("\n");
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

// TSV: no header, columns Name, Height, Width.
// Strips any prefix common to all artboard names so the spreadsheet
// has just the unique parts (e.g. "Driver - QP" instead of "Job_File_Driver - QP").
// Inserts a blank row between document sets.
function buildTSV(rows) {
    var names = [];
    for (var i = 0; i < rows.length; i++) names.push(rows[i].name);
    var commonPrefix = findCommonPrefix(names);

    var lines = [];
    var lastDocName = null;
    for (var j = 0; j < rows.length; j++) {
        var r = rows[j];
        if (lastDocName !== null && r.docName !== lastDocName) {
            lines.push(""); // blank row between document sets
        }
        var trimmedName = (commonPrefix.length > 0) ? r.name.substring(commonPrefix.length) : r.name;
        lines.push(trimmedName + "\t" + formatSize(r.finalH) + "\t" + formatSize(r.finalW));
        lastDocName = r.docName;
    }
    return lines.join("\n");
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
function calculatePanelEstimate(width, height, params) {
    var w1 = width + params.bleed * 2;
    var h1 = height + params.bleed * 2;
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
    var runWithWaste = runLength + params.rlWaste;
    var printedSqFt = ok ? (rollSize * runWithWaste) / 144 : 0;
    var installedSqFt = (width * height) / 144;

    return {
        installedSqFt: installedSqFt,
        printedSqFt: printedSqFt,
        rollRequired: rollRequired,
        runLength: runLength,
        runWithWaste: runWithWaste,
        rollSize: rollSize,
        ok: ok
    };
}

// Aggregate per-doc and overall estimate from the rows array.
function calculateOverallEstimate(rows, params) {
    var perDoc = {};         // docName -> { panels, installed, printed, errorPanels[] }
    var docOrder = [];
    var rollUsage = {};      // rollSize -> { count, sqft }
    var totals = { panels: 0, installed: 0, printed: 0, errorPanels: [] };

    for (var i = 0; i < rows.length; i++) {
        var r = rows[i];
        var est = calculatePanelEstimate(r.finalW, r.finalH, params);

        // attach for any later use
        r._estimate = est;

        if (!perDoc[r.docName]) {
            perDoc[r.docName] = {
                docName: r.docName,
                docBaseName: r.docBaseName,
                panels: 0,
                installed: 0,
                printed: 0,
                errorPanels: []
            };
            docOrder.push(r.docName);
        }
        var bucket = perDoc[r.docName];
        bucket.panels += 1;
        bucket.installed += est.installedSqFt;

        totals.panels += 1;
        totals.installed += est.installedSqFt;

        if (est.ok) {
            bucket.printed += est.printedSqFt;
            totals.printed += est.printedSqFt;

            var key = est.rollSize.toString();
            if (!rollUsage[key]) rollUsage[key] = { rollSize: est.rollSize, count: 0, sqft: 0, runInches: 0 };
            rollUsage[key].count += 1;
            rollUsage[key].sqft += est.printedSqFt;
            rollUsage[key].runInches += est.runWithWaste;
        } else {
            bucket.errorPanels.push(r.name);
            totals.errorPanels.push(r.name);
        }
    }

    // build sorted roll usage list
    var rollList = [];
    for (var k in rollUsage) {
        if (rollUsage.hasOwnProperty(k)) rollList.push(rollUsage[k]);
    }
    rollList.sort(function(a, b) { return a.rollSize - b.rollSize; });

    return {
        perDoc: perDoc,
        docOrder: docOrder,
        totals: totals,
        rollList: rollList,
        params: params
    };
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
    var t = estimate.totals;
    var lines = [];

    if (multiDoc) {
        for (var i = 0; i < estimate.docOrder.length; i++) {
            var d = estimate.perDoc[estimate.docOrder[i]];
            lines.push(d.docBaseName + ":  " + d.panels + " panels   Installed: " + formatSqFt(d.installed) + " sq ft   Printed: " + formatSqFt(d.printed) + " sq ft");
        }
        lines.push("");
    }

    lines.push("TOTAL:  " + t.panels + " panels");
    lines.push("Installed coverage: " + formatSqFt(t.installed) + " sq ft");
    lines.push("Printed coverage:   " + formatSqFt(t.printed) + " sq ft");

    var utilization = (t.printed > 0) ? (t.installed / t.printed) : 0;
    lines.push("Utilization: " + formatPct(utilization) + "    Waste: " + formatPct(1 - utilization));

    if (estimate.rollList.length > 0) {
        lines.push("");
        lines.push("Roll length needed:");
        for (var j = 0; j < estimate.rollList.length; j++) {
            var rl = estimate.rollList[j];
            lines.push("  " + rl.rollSize + "\" roll:  " + inchesToFeet(rl.runInches) + " ft   (" + formatSqFt(rl.sqft) + " sq ft)");
        }
    }

    if (t.errorPanels.length > 0) {
        lines.push("");
        lines.push("WARNING: " + t.errorPanels.length + " panel(s) too wide for any roll:");
        for (var e = 0; e < t.errorPanels.length; e++) {
            lines.push("  - " + t.errorPanels[e]);
        }
    }

    return lines.join("\n");
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

// Build email-friendly plain text — summary only (coverage + roll info).
// No per-section breakdown, no parameters block.
function buildEmailText(rows, estimate, multiDoc) {
    var t = estimate.totals;
    var info = getProjectInfo(estimate, multiDoc);
    var util = (t.printed > 0) ? (t.installed / t.printed) : 0;
    var lines = [];

    lines.push("Project: " + info.projectName);
    lines.push("");

    lines.push("TOTAL COVERAGE");
    lines.push("--------------");
    lines.push("Installed: " + formatSqFt(t.installed) + " sq ft");
    lines.push("Printed: " + formatSqFt(t.printed) + " sq ft");
    lines.push("Utilization: " + formatPct(util) + " (waste: " + formatPct(1 - util) + ")");
    lines.push("");

    if (estimate.rollList.length > 0) {
        lines.push("ROLL LENGTH NEEDED");
        lines.push("------------------");
        for (var j = 0; j < estimate.rollList.length; j++) {
            var rl = estimate.rollList[j];
            lines.push(rl.rollSize + "\" roll: " + inchesToFeet(rl.runInches) + " ft (" + formatSqFt(rl.sqft) + " sq ft)");
        }
    }

    if (t.errorPanels.length > 0) {
        lines.push("");
        lines.push("WARNING - Panels too wide for any stocked roll:");
        for (var e = 0; e < t.errorPanels.length; e++) {
            lines.push("- " + t.errorPanels[e]);
        }
    }

    return lines.join("\n");
}

// Escape special HTML characters
function escapeHtml(s) {
    if (s === undefined || s === null) return "";
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// Build rich-text HTML — summary only (coverage + roll info).
// Pasted into Outlook/Gmail/Apple Mail, this renders as a clean styled block.
function buildEmailHtml(rows, estimate, multiDoc) {
    var t = estimate.totals;
    var info = getProjectInfo(estimate, multiDoc);
    var util = (t.printed > 0) ? (t.installed / t.printed) : 0;

    var h3 = 'style="margin:20px 0 6px;font-size:13px;color:#444;text-transform:uppercase;letter-spacing:0.6px;font-weight:bold;border-bottom:1px solid #ddd;padding-bottom:4px;"';
    var tableStyle = 'style="border-collapse:collapse;font-size:13px;margin:0 0 8px 0;"';
    var thStyle = 'style="text-align:left;padding:6px 14px 6px 0;background:#f4f4f4;border-bottom:1px solid #ccc;font-weight:bold;"';
    var thRight = 'style="text-align:right;padding:6px 0 6px 14px;background:#f4f4f4;border-bottom:1px solid #ccc;font-weight:bold;"';
    var tdStyle = 'style="padding:5px 14px 5px 0;border-bottom:1px solid #eee;"';
    var tdRight = 'style="padding:5px 0 5px 14px;border-bottom:1px solid #eee;text-align:right;"';

    var html = '';
    html += '<div style="font-family:Arial,Helvetica,sans-serif;color:#1a1a1a;font-size:13px;line-height:1.5;">';

    html += '<div style="margin:0 0 12px;color:#222;font-size:14px;"><b>Project:</b> ' + escapeHtml(info.projectName) + '</div>';

    html += '<div ' + h3 + '>Total Coverage</div>';
    html += '<table ' + tableStyle + '>';
    html += '<tr><td ' + tdStyle + '>Installed</td><td ' + tdRight + '><b>' + formatSqFt(t.installed) + ' sq ft</b></td></tr>';
    html += '<tr><td ' + tdStyle + '>Printed</td><td ' + tdRight + '><b>' + formatSqFt(t.printed) + ' sq ft</b></td></tr>';
    html += '<tr><td ' + tdStyle + '>Utilization</td><td ' + tdRight + '>' + formatPct(util) + '</td></tr>';
    html += '<tr><td ' + tdStyle + '>Waste</td><td ' + tdRight + '>' + formatPct(1 - util) + '</td></tr>';
    html += '</table>';

    if (estimate.rollList.length > 0) {
        html += '<div ' + h3 + '>Roll Length Needed</div>';
        html += '<table ' + tableStyle + '>';
        html += '<tr>';
        html += '<th ' + thStyle + '>Roll</th>';
        html += '<th ' + thRight + '>Length</th>';
        html += '<th ' + thRight + '>Coverage</th>';
        html += '</tr>';
        for (var j = 0; j < estimate.rollList.length; j++) {
            var rl = estimate.rollList[j];
            html += '<tr>';
            html += '<td ' + tdStyle + '>' + rl.rollSize + '" roll</td>';
            html += '<td ' + tdRight + '><b>' + inchesToFeet(rl.runInches) + ' ft</b></td>';
            html += '<td ' + tdRight + '>' + formatSqFt(rl.sqft) + ' sq ft</td>';
            html += '</tr>';
        }
        html += '</table>';
    }

    if (t.errorPanels.length > 0) {
        html += '<div style="margin:20px 0 6px;font-size:13px;color:#c00;text-transform:uppercase;letter-spacing:0.6px;font-weight:bold;">Warning</div>';
        html += '<div style="color:#c00;margin-bottom:4px;">Panels too wide for any stocked roll:</div>';
        html += '<ul style="margin:0 0 12px 22px;color:#c00;">';
        for (var e = 0; e < t.errorPanels.length; e++) {
            html += '<li>' + escapeHtml(t.errorPanels[e]) + '</li>';
        }
        html += '</ul>';
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
        var ps = "";
        ps += "Add-Type -AssemblyName System.Windows.Forms\n";
        ps += "$html = Get-Content -Raw -Encoding UTF8 -Path '" + tempHtml.fsName.replace(/'/g, "''") + "'\n";
        ps += "$text = Get-Content -Raw -Encoding UTF8 -Path '" + tempText.fsName.replace(/'/g, "''") + "'\n";
        ps += "$do = New-Object System.Windows.Forms.DataObject\n";
        ps += "$do.SetText($text)\n";
        ps += "$do.SetText($html, [System.Windows.Forms.TextDataFormat]::Html)\n";
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
    var displayText = buildDisplayText(rows);

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

    // Roll-availability checkboxes — uncheck a size to force panels onto the
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

    // ---- Artboard list ----
    var listLabel = dlg.add("statictext", undefined, "Artboard sizes (x10, rounded up to 0.5\"):");
    var edit = dlg.add("edittext", undefined, displayText, {multiline: true, scrolling: true, readonly: false});
    edit.preferredSize = [600, 280];

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

    var copyTSVBtn = btnGroup.add("button", undefined, "Copy for Spreadsheet");
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

        // Only include checked roll sizes — defaults preserve sort order
        var rolls = [];
        for (var r = 0; r < rollCheckboxes.length; r++) {
            if (rollCheckboxes[r].checkbox.value) rolls.push(rollCheckboxes[r].width);
        }

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

    resetBtn.onClick = function() {
        bleedField.text = ESTIMATION_DEFAULTS.bleed.toString();
        wasteField.text = ESTIMATION_DEFAULTS.rlWaste.toString();
        marginField.text = ESTIMATION_DEFAULTS.pinchMargin.toString();
        for (var i = 0; i < rollCheckboxes.length; i++) {
            rollCheckboxes[i].checkbox.value = true;
        }
        refresh();
    };

    refresh();

    // ---- Button handlers ----
    copyTSVBtn.onClick = function() {
        var tsv = buildTSV(rows);
        var ok = setClipboardText(tsv);
        if (ok) {
            alert("Copied " + rows.length + " row(s) to clipboard.\nPaste into a spreadsheet — columns: Name, Height, Width.");
        } else {
            showCopyFallbackDialog(tsv);
        }
    };

    copyEmailBtn.onClick = function() {
        if (!currentEstimate) refresh();
        var emailText = buildEmailText(rows, currentEstimate, multiDoc);
        var ok = setClipboardText(emailText);
        if (ok) {
            alert("Email-ready estimate copied to clipboard.\nPaste into your email.");
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
