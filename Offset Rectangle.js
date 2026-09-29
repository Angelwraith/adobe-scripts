/*@METADATA{
  "name": "Offset Rectangle",
  "description": "Samples the bounds of the selected art (clipping-mask aware, same method as Smart Dimension Tool) and draws a rectangle that is that size plus or minus an offset on all sides. Example: 1.8 x 2 art with a 0.1 in offset makes a 2 x 2.2 rectangle. Negative offsets shrink the rectangle. Can make one rectangle around the whole selection or one per selected object.",
  "version": "1.0",
  "target": "illustrator",
  "tags": ["rectangle", "offset", "bounds", "measure", "utility"]
}@END_METADATA*/

#target illustrator

// ============================================================================
// SETTINGS
// ============================================================================
var PTS_PER_INCH = 72;
var ENV_OFFSET = "OffsetRect_offset";
var ENV_EACH = "OffsetRect_each";

// ============================================================================
// MAIN
// ============================================================================
(function () {
    try {
        main();
    } catch (e) {
        alert("Offset Rectangle error:\n" + e.message + (e.line ? "\nLine: " + e.line : ""));
    }
})();

function main() {
    if (app.documents.length === 0) {
        alert("Please open a document first.");
        return;
    }

    var doc = app.activeDocument;
    var sel = doc.selection;

    if (!sel || sel.length === 0 || sel.typename === "TextRange") {
        alert("Please select the artwork to measure.");
        return;
    }

    // Snapshot the selection -- sampling changes it.
    var items = [];
    for (var i = 0; i < sel.length; i++) {
        items.push(sel[i]);
    }

    // Sample bounds up front (menu commands can't run while a dialog is open).
    var unionBounds = getSimplifiedBounds(items);
    var eachBounds = [];
    if (items.length > 1) {
        for (var j = 0; j < items.length; j++) {
            eachBounds.push(getSimplifiedBounds([items[j]]));
        }
    }
    restoreSelection(items);

    var settings = showDialog(unionBounds, items.length);
    if (!settings) return;

    var offsetPts = settings.offset * PTS_PER_INCH;
    var boundsList = (settings.each && items.length > 1) ? eachBounds : [unionBounds];

    // Validate every rectangle before drawing anything.
    for (var k = 0; k < boundsList.length; k++) {
        var b = boundsList[k];
        var w = (b[2] - b[0]) + offsetPts * 2;
        var h = (b[1] - b[3]) + offsetPts * 2;
        if (w <= 0 || h <= 0) {
            alert("That offset makes a rectangle with no size (" +
                  fmt(w / PTS_PER_INCH) + " x " + fmt(h / PTS_PER_INCH) + " in).\n" +
                  "Use a smaller negative offset.");
            return;
        }
    }

    var layer = getTargetLayer(doc, items[0]);
    var created = [];
    for (var m = 0; m < boundsList.length; m++) {
        created.push(drawOffsetRect(layer, boundsList[m], offsetPts));
    }

    // Select the new rectangle(s).
    doc.selection = null;
    for (var n = 0; n < created.length; n++) {
        created[n].selected = true;
    }
    app.redraw();
}

// ============================================================================
// BOUNDS -- same approach as SmartDimensionTool.getSimplifiedBounds:
// temp artboard + "Fit Artboard to selected Art" so clipping masks and
// strokes are measured the way Illustrator sees them.
// ============================================================================
function getSimplifiedBounds(objs) {
    var doc = app.activeDocument;
    var originalArtboardIndex = -1;
    var artboardCountBefore = doc.artboards.length;

    try {
        originalArtboardIndex = doc.artboards.getActiveArtboardIndex();

        doc.artboards.add([0, 0, 100, -100]);
        var tempArtboardIndex = doc.artboards.length - 1;
        doc.artboards.setActiveArtboardIndex(tempArtboardIndex);

        doc.selection = null;
        for (var i = 0; i < objs.length; i++) {
            objs[i].selected = true;
        }

        app.executeMenuCommand("Fit Artboard to selected Art");
        app.redraw();

        var r = doc.artboards[tempArtboardIndex].artboardRect;
        var bounds = [r[0], r[1], r[2], r[3]];

        doc.artboards.remove(tempArtboardIndex);
        if (originalArtboardIndex >= 0 && originalArtboardIndex < doc.artboards.length) {
            doc.artboards.setActiveArtboardIndex(originalArtboardIndex);
        }
        return bounds;

    } catch (e) {
        try {
            if (doc.artboards.length > artboardCountBefore) {
                doc.artboards.remove(doc.artboards.length - 1);
            }
        } catch (e1) {}
        try {
            if (originalArtboardIndex >= 0 && originalArtboardIndex < doc.artboards.length) {
                doc.artboards.setActiveArtboardIndex(originalArtboardIndex);
            }
        } catch (e2) {}
        return unionGeometricBounds(objs);
    }
}

function unionGeometricBounds(objs) {
    var b = objs[0].geometricBounds;
    var out = [b[0], b[1], b[2], b[3]];
    for (var i = 1; i < objs.length; i++) {
        var g = objs[i].geometricBounds;
        if (g[0] < out[0]) out[0] = g[0];
        if (g[1] > out[1]) out[1] = g[1];
        if (g[2] > out[2]) out[2] = g[2];
        if (g[3] < out[3]) out[3] = g[3];
    }
    return out;
}

function restoreSelection(items) {
    try {
        app.activeDocument.selection = null;
        for (var i = 0; i < items.length; i++) {
            try { items[i].selected = true; } catch (e) {}
        }
    } catch (e2) {}
}

// ============================================================================
// DRAWING
// ============================================================================
function drawOffsetRect(layer, bounds, offsetPts) {
    var left = bounds[0] - offsetPts;
    var top = bounds[1] + offsetPts;
    var width = (bounds[2] - bounds[0]) + offsetPts * 2;
    var height = (bounds[1] - bounds[3]) + offsetPts * 2;

    var rect = layer.pathItems.rectangle(top, left, width, height);
    rect.filled = false;
    rect.stroked = true;
    rect.strokeWidth = 1;
    rect.strokeColor = blackColor();
    rect.name = "Offset Rect " + fmt(width / PTS_PER_INCH) + " x " + fmt(height / PTS_PER_INCH);
    return rect;
}

function blackColor() {
    if (app.activeDocument.documentColorSpace === DocumentColorSpace.CMYK) {
        var c = new CMYKColor();
        c.cyan = 0; c.magenta = 0; c.yellow = 0; c.black = 100;
        return c;
    }
    var r = new RGBColor();
    r.red = 0; r.green = 0; r.blue = 0;
    return r;
}

// Use the layer of the first selected item if it's editable, else the active layer.
function getTargetLayer(doc, item) {
    try {
        var lyr = item.layer;
        if (lyr && !lyr.locked && lyr.visible) return lyr;
    } catch (e) {}
    var active = doc.activeLayer;
    if (active.locked || !active.visible) {
        throw new Error("The target layer is locked or hidden. Unlock it and try again.");
    }
    return active;
}

// ============================================================================
// DIALOG
// ============================================================================
function showDialog(bounds, itemCount) {
    var artW = (bounds[2] - bounds[0]) / PTS_PER_INCH;
    var artH = (bounds[1] - bounds[3]) / PTS_PER_INCH;

    var lastOffset = $.getenv(ENV_OFFSET);
    if (lastOffset === null || lastOffset === "" || isNaN(parseFloat(lastOffset))) lastOffset = "0.1";
    var lastEach = $.getenv(ENV_EACH) === "1";

    var dlg = new Window("dialog", "Offset Rectangle");
    dlg.alignChildren = "fill";

    var infoPanel = dlg.add("panel", undefined, "Sampled Art");
    infoPanel.alignChildren = "left";
    infoPanel.add("statictext", undefined,
        "Selection: " + fmt(artW) + " x " + fmt(artH) + " in" +
        (itemCount > 1 ? "  (" + itemCount + " objects)" : ""));

    var offPanel = dlg.add("panel", undefined, "Offset (all sides)");
    offPanel.alignChildren = "left";
    var row = offPanel.add("group");
    row.add("statictext", undefined, "Offset:");
    var offInput = row.add("edittext", undefined, lastOffset);
    offInput.characters = 8;
    row.add("statictext", undefined, "in  (negative = smaller)");

    var eachChk = offPanel.add("checkbox", undefined, "One rectangle per selected object");
    eachChk.value = lastEach;
    eachChk.enabled = itemCount > 1;
    if (itemCount <= 1) eachChk.value = false;

    var resultTxt = offPanel.add("statictext", undefined, "", { multiline: false });
    resultTxt.characters = 40;

    function readOffset() {
        var v = parseFloat(offInput.text);
        return isNaN(v) ? null : v;
    }

    function refresh() {
        var v = readOffset();
        if (v === null) {
            resultTxt.text = "Enter a number.";
            return;
        }
        if (eachChk.value) {
            resultTxt.text = "Each rectangle: object size + " + fmt(v * 2) + " in";
        } else {
            resultTxt.text = "New rectangle: " + fmt(artW + v * 2) + " x " + fmt(artH + v * 2) + " in";
        }
    }

    offInput.onChanging = refresh;
    eachChk.onClick = refresh;
    refresh();

    var btns = dlg.add("group");
    btns.alignment = "right";
    btns.add("button", undefined, "Cancel", { name: "cancel" });
    var okBtn = btns.add("button", undefined, "OK", { name: "ok" });
    okBtn.onClick = function () {
        if (readOffset() === null) {
            alert("Offset must be a number (inches).");
            return;
        }
        dlg.close(1);
    };

    offInput.active = true;
    if (dlg.show() !== 1) return null;

    var offset = readOffset();
    $.setenv(ENV_OFFSET, String(offset));
    $.setenv(ENV_EACH, eachChk.value ? "1" : "0");
    return { offset: offset, each: eachChk.value };
}

// ============================================================================
// HELPERS
// ============================================================================
function fmt(n) {
    var r = Math.round(n * 1000) / 1000;
    return String(r);
}
