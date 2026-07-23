/*
@METADATA
{
  "name": "Scale To Page",
  "description": "Copies the selected 1/10 scale art onto the active artboard at a chosen percentage of its current size, keeping the relative layout, then writes a matching \"Scale 1:N\" label at the bottom of the page. Starting scale is assumed to be 1:10, so 50% -> 1:20, 200% -> 1:5, etc. You can enter either a percentage (50%, 2x) or a target ratio (1:12) directly; the resulting scale and equivalent percentage update live as you type, so decimal scales are easy to spot and avoid. The label is written in the format the Smart Dimension Tool auto-detects, so dimensions come out accurate with no extra setup.",
  "version": "1.1",
  "target": "illustrator",
  "tags": ["scale", "copy", "layout", "processor"]
}
@END_METADATA
*/

#target illustrator

(function () {
    'use strict';

    if (app.documents.length === 0) {
        alert("Please open a document first.");
        return;
    }

    var doc = app.activeDocument;
    var sel = doc.selection;

    if (!sel || sel.length === 0) {
        alert("Select the 1/10 scale art you want to copy onto the page, then run the script.");
        return;
    }

    // --- Snapshot the current selection (running the dialog can clear it) ---
    var sourceItems = [];
    for (var i = 0; i < sel.length; i++) {
        sourceItems.push(sel[i]);
    }

    // ------------------------------------------------------------------
    // Helpers
    // ------------------------------------------------------------------

    // Look up a character style by name (matches SmartDimensionTool's DimStyle usage)
    function getCharacterStyle(styleName) {
        try {
            for (var i = 0; i < doc.characterStyles.length; i++) {
                if (doc.characterStyles[i].name === styleName) {
                    return doc.characterStyles[i];
                }
            }
        } catch (e) {}
        return null;
    }

    // Trim trailing zeros / format a number cleanly for the scale label.
    function fmtNum(n) {
        var r = Math.round(n * 100) / 100;      // 2 decimals max
        if (Math.abs(r - Math.round(r)) < 1e-9) {
            return String(Math.round(r));       // whole number
        }
        var s = r.toFixed(2);
        s = s.replace(/0+$/, "").replace(/\.$/, "");
        return s;
    }

    // Parse a percentage input. Accepts "50", "50%", "2x", "x2".
    // Returns a percent number (e.g. 50, 200) or NaN.
    function parsePercent(raw) {
        if (raw === null) return NaN;
        var s = String(raw).replace(/^\s+|\s+$/g, "").toLowerCase();
        if (s.length === 0) return NaN;
        var mult = false;
        if (s.charAt(0) === "x") { mult = true; s = s.substring(1); }
        else if (s.charAt(s.length - 1) === "x") { mult = true; s = s.substring(0, s.length - 1); }
        s = s.replace(/%/g, "").replace(/^\s+|\s+$/g, "");
        var v = parseFloat(s);
        if (isNaN(v)) return NaN;
        return mult ? v * 100 : v;
    }

    // Parse either a percentage ("50", "50%", "2x") OR a target scale ratio
    // ("1:12", "2:24"). Returns { pct, factor, denominator } or null.
    //   - percentage: factor = pct/100, on-page denominator = 10*100/pct
    //   - ratio a:b : on-page denominator = b/a, factor = 10/denominator
    function parseInput(raw) {
        if (raw === null) return null;
        var s = String(raw).replace(/^\s+|\s+$/g, "");
        if (s.length === 0) return null;

        if (s.indexOf(":") !== -1) {
            var parts = s.split(":");
            var a = parseFloat(parts[0]);
            var b = parseFloat(parts[1]);
            if (isNaN(a) || isNaN(b) || a <= 0 || b <= 0) return null;
            var denom = b / a;                              // on-page scale 1:denom
            var factorR = BASE_DENOMINATOR / denom;         // multiplier vs current size
            return { pct: factorR * 100, factor: factorR, denominator: denom };
        }

        var pct = parsePercent(s);
        if (isNaN(pct) || pct <= 0) return null;
        return { pct: pct, factor: pct / 100, denominator: BASE_DENOMINATOR * 100 / pct };
    }

    // Collective geometric bounds of an array of page items -> [L, T, R, B]
    function collectiveBounds(items) {
        var minL = Infinity, maxT = -Infinity, maxR = -Infinity, minB = Infinity;
        for (var k = 0; k < items.length; k++) {
            var b = items[k].geometricBounds; // [L, T, R, B]
            if (b[0] < minL) minL = b[0];
            if (b[1] > maxT) maxT = b[1];
            if (b[2] > maxR) maxR = b[2];
            if (b[3] < minB) minB = b[3];
        }
        return [minL, maxT, maxR, minB];
    }

    // ------------------------------------------------------------------
    // Dialog
    // ------------------------------------------------------------------

    var dlg = new Window("dialog", "Scale To Page");
    dlg.orientation = "column";
    dlg.alignChildren = "fill";
    dlg.margins = 16;
    dlg.spacing = 10;

    var BASE_DENOMINATOR = 10; // art is always extracted at 1:10 scale

    dlg.add("statictext", undefined, "Selected art: " + sourceItems.length + " object(s)   (starting scale 1:10)");

    // Resize amount: a percentage OR a target scale ratio
    var pctGroup = dlg.add("group");
    pctGroup.add("statictext", undefined, "Resize to:");
    var pctInput = pctGroup.add("edittext", undefined, "100%");
    pctInput.characters = 8;
    pctGroup.add("statictext", undefined, "% of current size, or a target ratio like 1:12");

    // Live preview of the resulting scale
    var previewText = dlg.add("statictext", undefined, "New scale on page:  1:10   (100% of current)");
    previewText.preferredSize.width = 360; // room for longer messages
    previewText.graphics.font = ScriptUI.newFont(previewText.graphics.font.name, "BOLD", 13);

    // Options
    var optPanel = dlg.add("panel", undefined, "Options");
    optPanel.orientation = "column";
    optPanel.alignChildren = "left";
    optPanel.margins = 12;
    optPanel.add("statictext", undefined, "Copies are always placed on the ACTIVE artboard.");
    var cbLabel = optPanel.add("checkbox", undefined, "Add \"Scale 1:N\" label (matches Smart Dimension Tool exactly)");
    cbLabel.value = true;

    function refreshPreview() {
        var r = parseInput(pctInput.text);
        if (!r) {
            previewText.text = "New scale on page:  --";
            return;
        }
        var txt = "New scale on page:  1:" + fmtNum(r.denominator) +
                  "   (" + fmtNum(r.pct) + "% of current)";
        // Flag non-whole ratios so decimals are easy to avoid
        if (Math.abs(r.denominator - Math.round(r.denominator)) > 1e-9) {
            txt += "   <-- has decimals";
        }
        previewText.text = txt;
    }
    pctInput.onChanging = refreshPreview;
    refreshPreview();

    // Buttons
    var btnGroup = dlg.add("group");
    btnGroup.alignment = "right";
    var cancelBtn = btnGroup.add("button", undefined, "Cancel", { name: "cancel" });
    var okBtn = btnGroup.add("button", undefined, "OK", { name: "ok" });

    if (dlg.show() !== 1) {
        return; // cancelled
    }

    // ------------------------------------------------------------------
    // Validate inputs
    // ------------------------------------------------------------------

    var parsed = parseInput(pctInput.text);

    if (!parsed) {
        alert("Please enter a valid resize amount.\n\nExamples:\n  50%    (half of current size)\n  2x     (double)\n  1:12   (target on-page scale)");
        return;
    }

    var factor = parsed.factor;             // scale multiplier applied to the art
    var pct = parsed.pct;                   // percentage of current size
    var denominator = parsed.denominator;   // resulting scale ratio 1:denominator
    var scaleString = "1:" + fmtNum(denominator);
    var doLabel = cbLabel.value;

    // ------------------------------------------------------------------
    // 1. Duplicate the source art
    // ------------------------------------------------------------------

    var copies = [];
    try {
        for (var d = 0; d < sourceItems.length; d++) {
            copies.push(sourceItems[d].duplicate());
        }
    } catch (e) {
        alert("Could not duplicate the selection.\n" + e);
        return;
    }

    // ------------------------------------------------------------------
    // 2. Scale the copies as a unit (preserves relative layout)
    //    Anchor everything to the collective top-left, scale about that
    //    point so the whole arrangement grows/shrinks together.
    // ------------------------------------------------------------------

    var pre = collectiveBounds(copies);
    var anchorX = pre[0]; // collective left
    var anchorY = pre[1]; // collective top

    for (var c = 0; c < copies.length; c++) {
        var item = copies[c];
        var b = item.geometricBounds;   // [L, T, R, B]
        var objL = b[0];
        var objT = b[1];

        // Scale the object about its own top-left (keeps that point fixed)
        item.resize(
            factor * 100, factor * 100,
            true,  // changePositions
            true,  // changeFillPatterns
            true,  // changeFillGradients
            true,  // changeStrokePattern
            true,  // changeLineWidths
            Transformation.TOPLEFT
        );

        // Move its top-left so its offset from the anchor is scaled too
        var newL = anchorX + factor * (objL - anchorX);
        var newT = anchorY + factor * (objT - anchorY);
        item.translate(newL - objL, newT - objT);
    }

    // ------------------------------------------------------------------
    // 3. Place the scaled copies on the ACTIVE artboard (always centered
    //    there, so there is never any guessing about where they land)
    // ------------------------------------------------------------------

    var abIndex = doc.artboards.getActiveArtboardIndex();
    var abRect = doc.artboards[abIndex].artboardRect; // [L, T, R, B]
    var abCenterX = (abRect[0] + abRect[2]) / 2;
    var abCenterY = (abRect[1] + abRect[3]) / 2;

    var post = collectiveBounds(copies);
    var curCenterX = (post[0] + post[2]) / 2;
    var curCenterY = (post[1] + post[3]) / 2;
    var dx = abCenterX - curCenterX;
    var dy = abCenterY - curCenterY;
    for (var m = 0; m < copies.length; m++) {
        copies[m].translate(dx, dy);
    }

    // ------------------------------------------------------------------
    // 4. Write / replace the "Scale 1:N" label.
    //    Built to match SmartDimensionTool's "On Proof" setting exactly:
    //    same "Scale " + value contents, DimStyle character style, 7pt,
    //    centered, positioned 4" from the artboard's left edge and 6.9816"
    //    down from its top. This is precisely where the dimension tool's
    //    "On Proof" option would place it, so it's already set.
    // ------------------------------------------------------------------

    if (doLabel) {
        // Remove any existing "Scale ..." label already sitting on this artboard
        try {
            for (var t = doc.textFrames.length - 1; t >= 0; t--) {
                var tf = doc.textFrames[t];
                var contents = "";
                try { contents = tf.contents; } catch (eC) { contents = ""; }
                if (contents.toLowerCase().indexOf("scale ") !== 0) continue;

                var tb = tf.geometricBounds; // [L, T, R, B]
                var cx = (tb[0] + tb[2]) / 2;
                var cy = (tb[1] + tb[3]) / 2;
                // Inside the active artboard?
                if (cx >= abRect[0] && cx <= abRect[2] &&
                    cy <= abRect[1] && cy >= abRect[3]) {
                    tf.remove();
                }
            }
        } catch (eRem) { /* non-fatal */ }

        // Pick a usable layer for the label (active layer, fall back if locked/hidden)
        var labelLayer = doc.activeLayer;
        try {
            if (labelLayer.locked || !labelLayer.visible) {
                for (var L = 0; L < doc.layers.length; L++) {
                    if (!doc.layers[L].locked && doc.layers[L].visible) {
                        labelLayer = doc.layers[L];
                        break;
                    }
                }
            }
        } catch (eLay) {}

        try {
            // "On Proof" position: fixed relative to the active artboard.
            //   x = 4"   from the artboard's left edge
            //   y = 6.9816" down from the artboard's top edge
            var labelX = abRect[0] + (4 * 72);
            var labelY = abRect[1] - (6.9816 * 72);

            var label = labelLayer.textFrames.add();
            label.contents = "Scale " + scaleString;

            // Apply DimStyle character style if it exists (matches dimension tool)
            var characterStyle = getCharacterStyle("DimStyle");
            if (characterStyle) {
                label.textRange.characterAttributes.characterStyle = characterStyle;
            }

            label.textRange.characterAttributes.size = 7;
            label.textRange.paragraphAttributes.justification = Justification.CENTER;

            label.top = labelY;
            label.left = labelX - (label.width / 2);
        } catch (eLbl) {
            alert("Copies placed, but the scale label could not be added.\n" + eLbl);
        }
    }

    // ------------------------------------------------------------------
    // 5. Leave the new copies selected for the next step (dimension tool)
    // ------------------------------------------------------------------

    try {
        doc.selection = null;
        for (var s = 0; s < copies.length; s++) {
            copies[s].selected = true;
        }
    } catch (eSel) { /* non-fatal */ }

    app.redraw();

})();
