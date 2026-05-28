/*@METADATA{
  "name": "Label Artboards",
  "description": "Adds a text label below each artboard using the artboard name. Skips artboards named 'proof'. Labels go on a dedicated 'Artboard Labels' layer.",
  "version": "1.0",
  "target": "illustrator",
  "tags": ["artboard", "label", "text", "name"]
}@END_METADATA*/

#target illustrator

// ============================================================================
// CONSTANTS
// ============================================================================
var LABEL_DEFAULTS = {
    fontSize: 18,           // points
    offset: 18,             // points below the artboard's bottom edge
    layerName: "Artboard Labels",
    skipProof: true         // skip artboards named "proof" (case-insensitive)
};

// ============================================================================
// MAIN ENTRY POINT
// ============================================================================
try {
    if (app.documents.length === 0) {
        alert("Please open a document first.");
    } else {
        labelArtboards();
    }
} catch (e) {
    alert("Error: " + e.toString() + "\n\nLine: " + e.line);
}

// ============================================================================
// MAIN FLOW
// ============================================================================
function labelArtboards() {
    var doc = app.activeDocument;

    if (doc.artboards.length === 0) {
        alert("This document has no artboards.");
        return;
    }

    var opts = showOptionsDialog();
    if (!opts) return;

    var result = addLabelsToArtboards(doc, opts);

    var msg = "Added " + result.added + " label(s) on layer \"" + opts.layerName + "\".";
    if (result.skipped > 0) {
        msg += "\nSkipped " + result.skipped + " artboard(s) named \"proof\".";
    }
    if (result.errors.length > 0) {
        msg += "\n\nErrors on " + result.errors.length + " artboard(s):\n" + result.errors.join("\n");
    }
    alert(msg);
}

// ============================================================================
// OPTIONS DIALOG
// ============================================================================
function showOptionsDialog() {
    var dialog = new Window("dialog", "Label Artboards");
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.spacing = 15;
    dialog.margins = 20;
    dialog.preferredSize.width = 320;

    var titleText = dialog.add("statictext", undefined, "Add a text label below each artboard:");
    titleText.graphics.font = ScriptUI.newFont("dialog", "Bold", 12);

    var fieldsPanel = dialog.add("panel");
    fieldsPanel.orientation = "column";
    fieldsPanel.alignChildren = "left";
    fieldsPanel.margins = 12;
    fieldsPanel.spacing = 10;

    var row1 = fieldsPanel.add("group");
    row1.orientation = "row";
    row1.spacing = 10;
    row1.add("statictext", undefined, "Font size (pt):");
    var sizeField = row1.add("edittext", undefined, LABEL_DEFAULTS.fontSize.toString());
    sizeField.preferredSize.width = 60;

    var row2 = fieldsPanel.add("group");
    row2.orientation = "row";
    row2.spacing = 10;
    row2.add("statictext", undefined, "Offset below (pt):");
    var offsetField = row2.add("edittext", undefined, LABEL_DEFAULTS.offset.toString());
    offsetField.preferredSize.width = 60;

    var row3 = fieldsPanel.add("group");
    row3.orientation = "row";
    row3.spacing = 10;
    row3.add("statictext", undefined, "Layer name:");
    var layerField = row3.add("edittext", undefined, LABEL_DEFAULTS.layerName);
    layerField.preferredSize.width = 160;

    var skipProofCheckbox = fieldsPanel.add("checkbox", undefined, "Skip artboards named \"proof\"");
    skipProofCheckbox.value = LABEL_DEFAULTS.skipProof;

    var btnGroup = dialog.add("group");
    btnGroup.alignment = "center";
    btnGroup.spacing = 10;
    var cancelBtn = btnGroup.add("button", undefined, "Cancel", {name: "cancel"});
    var okBtn = btnGroup.add("button", undefined, "Add Labels", {name: "ok"});

    var result = null;
    cancelBtn.onClick = function() { result = null; dialog.close(); };
    okBtn.onClick = function() {
        var fs = parseFloat(sizeField.text);
        var off = parseFloat(offsetField.text);
        if (isNaN(fs) || fs <= 0) {
            alert("Font size must be a positive number.");
            return;
        }
        if (isNaN(off)) {
            alert("Offset must be a number.");
            return;
        }
        var ln = layerField.text;
        if (!ln || ln.replace(/\s/g, "") === "") {
            alert("Layer name cannot be empty.");
            return;
        }
        result = {
            fontSize: fs,
            offset: off,
            layerName: ln,
            skipProof: skipProofCheckbox.value
        };
        dialog.close();
    };

    dialog.show();
    return result;
}

// ============================================================================
// LAYER HELPER
// ============================================================================
// Find or create the labels layer. Ensures it's unlocked and visible so
// new text frames can be created on it.
function getOrCreateLayer(doc, name) {
    var layer = null;
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === name) {
            layer = doc.layers[i];
            break;
        }
    }
    if (!layer) {
        layer = doc.layers.add();
        layer.name = name;
    }
    try { layer.locked = false; } catch (e) {}
    try { layer.visible = true; } catch (e) {}
    return layer;
}

// ============================================================================
// ADD LABELS
// ============================================================================
function addLabelsToArtboards(doc, opts) {
    var labelsLayer = getOrCreateLayer(doc, opts.layerName);

    // Text color appropriate to the doc's color mode
    var textColor;
    if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
        textColor = new CMYKColor();
        textColor.cyan = 0; textColor.magenta = 0; textColor.yellow = 0; textColor.black = 100;
    } else {
        textColor = new RGBColor();
        textColor.red = 0; textColor.green = 0; textColor.blue = 0;
    }

    var added = 0;
    var skipped = 0;
    var errors = [];

    for (var i = 0; i < doc.artboards.length; i++) {
        var ab = doc.artboards[i];

        if (opts.skipProof && ab.name.toLowerCase() === "proof") {
            skipped++;
            continue;
        }

        try {
            var rect = ab.artboardRect; // [left, top, right, bottom] in points
            var centerX = (rect[0] + rect[2]) / 2;
            var bottomY = rect[3];

            // Create the text frame at the artboard's horizontal center,
            // offset below the bottom edge. Position is set after creation
            // since geometricBounds isn't valid until the frame has content.
            var tf = labelsLayer.textFrames.add();
            tf.contents = ab.name;

            // Apply font size to all characters
            try {
                tf.textRange.characterAttributes.size = opts.fontSize;
                tf.textRange.characterAttributes.fillColor = textColor;
            } catch (e) {
                // some Illustrator builds need per-character access
                for (var c = 0; c < tf.characters.length; c++) {
                    tf.characters[c].characterAttributes.size = opts.fontSize;
                    tf.characters[c].characterAttributes.fillColor = textColor;
                }
            }

            // Force Illustrator to compute the frame's bounds, then center
            // it horizontally and place its top edge `offset` below artboard
            app.redraw();
            var tb = tf.geometricBounds; // [left, top, right, bottom]
            var tWidth = tb[2] - tb[0];
            var currentLeft = tb[0];
            var currentTop = tb[1];

            var targetLeft = centerX - (tWidth / 2);
            var targetTop = bottomY - opts.offset;

            tf.translate(targetLeft - currentLeft, targetTop - currentTop);

            added++;
        } catch (e) {
            errors.push(ab.name + ": " + e.toString());
        }
    }

    app.redraw();

    return { added: added, skipped: skipped, errors: errors };
}
