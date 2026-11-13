/*@METADATA{
  "name": "Round Up Artboard Sizes",
  "description": "Increases the width and height of chosen artboards to the next whole inch. Defaults to all artboards, with All/None buttons.",
  "version": "1.1",
  "target": "illustrator",
  "tags": ["artboard", "resize", "round", "size"]
}@END_METADATA*/

#target illustrator

// ============================================================================
// CONSTANTS
// ============================================================================
var PT_PER_IN = 72;
var UNITS_PER_IN = 72;         // script units per inch; adjusted for Large Canvas docs
var EPSILON_IN = 0.001;          // sizes within this of a whole inch count as whole
var MAX_ARTBOARD_PT = 16383;     // Illustrator artboard size limit (~227.5 in)
var ANCHORS = ["Center", "Top Left", "Top Center", "Bottom Left", "Bottom Center"];

// ============================================================================
// MAIN ENTRY POINT
// ============================================================================
try {
    if (app.documents.length === 0) {
        alert("Please open a document first.");
    } else {
        roundUpArtboards();
    }
} catch (e) {
    alert("Error: " + e.toString() + "\n\nLine: " + e.line);
}

// ============================================================================
// MAIN FLOW
// ============================================================================
function roundUpArtboards() {
    var doc = app.activeDocument;

    if (doc.artboards.length === 0) {
        alert("This document has no artboards.");
        return;
    }

    // Large Canvas documents report coordinates at 1/scaleFactor (usually 1/10)
    var sf = 1;
    try { if (doc.scaleFactor && doc.scaleFactor > 0) sf = doc.scaleFactor; } catch (e) {}
    UNITS_PER_IN = PT_PER_IN / sf;

    var info = collectArtboardInfo(doc);
    var opts = showOptionsDialog(info);
    if (!opts) return;

    var changed = 0;
    var unchanged = 0;
    var errors = [];

    for (var i = 0; i < opts.indices.length; i++) {
        var a = info[opts.indices[i]];
        if (!a.needsChange) { unchanged++; continue; }
        try {
            doc.artboards[a.index].artboardRect = buildRect(a, opts.anchor);
            changed++;
        } catch (e) {
            errors.push((a.index + 1) + ": " + a.name + " - " + e.toString());
        }
    }

    var msg = "Resized " + changed + " artboard(s) to the next whole inch.";
    if (unchanged > 0) msg += "\n" + unchanged + " artboard(s) were already whole inches.";
    if (errors.length > 0) msg += "\n\nErrors on " + errors.length + " artboard(s):\n" + errors.join("\n");
    alert(msg);
}

// ============================================================================
// SIZE HELPERS
// ============================================================================
function ceilInches(sizeIn) {
    var r = Math.round(sizeIn);
    if (Math.abs(sizeIn - r) < EPSILON_IN) return r;   // already whole
    return Math.ceil(sizeIn);
}

function fmt(n) {
    var s = (Math.round(n * 1000) / 1000).toString();
    return s;
}

function collectArtboardInfo(doc) {
    var list = [];
    for (var i = 0; i < doc.artboards.length; i++) {
        var ab = doc.artboards[i];
        var r = ab.artboardRect; // [left, top, right, bottom]
        var wIn = (r[2] - r[0]) / UNITS_PER_IN;
        var hIn = (r[1] - r[3]) / UNITS_PER_IN;
        var newW = ceilInches(wIn);
        var newH = ceilInches(hIn);
        if (newW * UNITS_PER_IN > MAX_ARTBOARD_PT) newW = wIn;
        if (newH * UNITS_PER_IN > MAX_ARTBOARD_PT) newH = hIn;
        list.push({
            index: i,
            name: ab.name,
            rect: r,
            wIn: wIn,
            hIn: hIn,
            newW: newW,
            newH: newH,
            needsChange: (Math.abs(newW - wIn) >= EPSILON_IN || Math.abs(newH - hIn) >= EPSILON_IN)
        });
    }
    return list;
}

// Build the new artboardRect, growing away from the chosen anchor point
function buildRect(a, anchor) {
    var r = a.rect;
    var w = a.newW * UNITS_PER_IN;
    var h = a.newH * UNITS_PER_IN;
    var left, top;

    var cx = (r[0] + r[2]) / 2;
    var cy = (r[1] + r[3]) / 2;

    switch (anchor) {
        case "Top Left":      left = r[0];      top = r[1];      break;
        case "Top Center":    left = cx - w / 2; top = r[1];     break;
        case "Bottom Left":   left = r[0];      top = r[3] + h;  break;
        case "Bottom Center": left = cx - w / 2; top = r[3] + h; break;
        default:              left = cx - w / 2; top = cy + h / 2;
    }
    return [left, top, left + w, top - h];
}

// ============================================================================
// OPTIONS DIALOG
// ============================================================================
function showOptionsDialog(info) {
    var dialog = new Window("dialog", "Round Up Artboard Sizes");
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.spacing = 12;
    dialog.margins = 20;

    var titleText = dialog.add("statictext", undefined, "Round artboard sizes up to the next whole inch:");
    titleText.graphics.font = ScriptUI.newFont("dialog", "Bold", 12);

    var listPanel = dialog.add("panel", undefined, "Artboards");
    listPanel.orientation = "column";
    listPanel.alignChildren = "fill";
    listPanel.margins = 12;
    listPanel.spacing = 8;

    var list = listPanel.add("listbox", undefined, [], {multiselect: true});
    list.preferredSize = [460, 260];
    for (var i = 0; i < info.length; i++) {
        var a = info[i];
        var label = (a.index + 1) + ": " + a.name + "   " + fmt(a.wIn) + " x " + fmt(a.hIn) + " in";
        label += a.needsChange ? "  ->  " + fmt(a.newW) + " x " + fmt(a.newH) + " in" : "  (already whole)";
        list.add("item", label);
    }
    for (var s = 0; s < list.items.length; s++) list.items[s].selected = true; // default: all

    var selBtns = listPanel.add("group");
    selBtns.orientation = "row";
    selBtns.alignment = "left";
    var allBtn = selBtns.add("button", undefined, "All");
    var noneBtn = selBtns.add("button", undefined, "None");
    var countText = selBtns.add("statictext", undefined, "");
    countText.preferredSize.width = 160;

    function updateCount() {
        var n = 0;
        for (var k = 0; k < list.items.length; k++) if (list.items[k].selected) n++;
        countText.text = n + " of " + list.items.length + " selected";
    }
    allBtn.onClick = function() {
        for (var k = 0; k < list.items.length; k++) list.items[k].selected = true;
        updateCount();
    };
    noneBtn.onClick = function() {
        list.selection = null;
        updateCount();
    };
    list.onChange = updateCount;
    updateCount();

    var anchorRow = dialog.add("group");
    anchorRow.orientation = "row";
    anchorRow.add("statictext", undefined, "Grow from:");
    var anchorDrop = anchorRow.add("dropdownlist", undefined, ANCHORS);
    anchorDrop.selection = 0;

    var btnGroup = dialog.add("group");
    btnGroup.alignment = "center";
    btnGroup.spacing = 10;
    var cancelBtn = btnGroup.add("button", undefined, "Cancel", {name: "cancel"});
    var okBtn = btnGroup.add("button", undefined, "Round Up", {name: "ok"});

    var result = null;
    cancelBtn.onClick = function() { result = null; dialog.close(); };
    okBtn.onClick = function() {
        var picked = [];
        for (var k = 0; k < list.items.length; k++) if (list.items[k].selected) picked.push(k);
        if (picked.length === 0) {
            alert("Select at least one artboard.");
            return;
        }
        result = { indices: picked, anchor: anchorDrop.selection.text };
        dialog.close();
    };

    dialog.show();
    return result;
}
