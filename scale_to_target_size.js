/*
@METADATA
{
  "name": "Scale Selection to Target Size",
  "description": "Select a key object, enter target dimensions, and scale everything proportionally. Also supports ratio conversion (e.g., 3/8\"=1'-0\" -> 1:10).",
  "version": "1.1",
  "target": "illustrator",
  "tags": ["scale", "size", "processors"]
}
@END_METADATA
*/
function main() {
    if (app.documents.length === 0) {
        alert("Please open a document first.");
        return;
    }

    var doc = app.activeDocument;
    var sel = doc.selection;

    if (sel.length === 0) {
        alert("Please select at least one object.\n\nFor size-based modes, the first selected object is used as the key object.\nFor ratio conversion mode, the scale is applied uniformly to the entire selection.");
        return;
    }

    // The first selected object is the key object (used for size-based scaling)
    var keyObject = sel[0];

    // Get current dimensions of key object
    var bounds = keyObject.geometricBounds; // [left, top, right, bottom]
    var currentWidth = bounds[2] - bounds[0];
    var currentHeight = bounds[1] - bounds[3];

    // Convert from points to inches (72 points = 1 inch)
    var currentWidthInches = currentWidth / 72;
    var currentHeightInches = currentHeight / 72;

    // Create dialog
    var dialog = new Window("dialog", "Scale to Target Size");
    dialog.orientation = "column";
    dialog.alignChildren = "fill";

    // Current size display
    dialog.add("statictext", undefined, "Current key object size:");
    dialog.add("statictext", undefined, "Width: " + currentWidthInches.toFixed(4) + '"  Height: ' + currentHeightInches.toFixed(4) + '"');
    dialog.add("statictext", undefined, ""); // Spacer

    // ----- Target size inputs (used by Width/Height/Exact/Fit modes) -----
    var sizePanel = dialog.add("group");
    sizePanel.orientation = "column";
    sizePanel.alignChildren = "left";

    var widthGroup = sizePanel.add("group");
    widthGroup.add("statictext", undefined, "Target Width (inches):");
    var widthInput = widthGroup.add("edittext", undefined, "");
    widthInput.characters = 10;

    var heightGroup = sizePanel.add("group");
    heightGroup.add("statictext", undefined, "Target Height (inches):");
    var heightInput = heightGroup.add("edittext", undefined, "");
    heightInput.characters = 10;

    // ----- Ratio conversion inputs (used by Ratio mode) -----
    var ratioPanel = dialog.add("panel", undefined, "Ratio Conversion");
    ratioPanel.orientation = "column";
    ratioPanel.alignChildren = "left";
    ratioPanel.margins = 10;
    ratioPanel.visible = false;

    var sourceGroup = ratioPanel.add("group");
    sourceGroup.add("statictext", undefined, "Source Scale:");
    var sourceInput = sourceGroup.add("edittext", undefined, "");
    sourceInput.characters = 18;
    sourceInput.helpTip = "Examples: 3/8\" = 1'-0\"   |   1/4\" = 1'   |   1:32   |   1 1/2\" = 1'";

    var sourceHint = ratioPanel.add("statictext", undefined, "(e.g.  3/8\" = 1'-0\"   or   1:32)");

    var targetGroup = ratioPanel.add("group");
    targetGroup.add("statictext", undefined, "Target Scale:");
    var targetInput = targetGroup.add("edittext", undefined, "1:10");
    targetInput.characters = 18;
    targetInput.helpTip = "Examples: 1:10   |   1:20   |   1\" = 10\"   |   1\" = 1'";

    var targetHint = ratioPanel.add("statictext", undefined, "(e.g.  1:10   default)");

    var ratioPreview = ratioPanel.add("statictext", undefined, "                                                                            ");
    ratioPreview.preferredSize.width = 360;

    dialog.add("statictext", undefined, ""); // Spacer

    // ----- Scale method radio buttons -----
    var scaleGroup = dialog.add("panel", undefined, "Scale Method");
    scaleGroup.orientation = "column";
    scaleGroup.alignChildren = "left";
    var radioWidth = scaleGroup.add("radiobutton", undefined, "Scale by Width (maintain aspect ratio)");
    var radioHeight = scaleGroup.add("radiobutton", undefined, "Scale by Height (maintain aspect ratio)");
    var radioBoth = scaleGroup.add("radiobutton", undefined, "Scale to exact dimensions (may distort)");
    var radioFit = scaleGroup.add("radiobutton", undefined, "Fit proportionally (fit inside dimensions)");
    var radioRatio = scaleGroup.add("radiobutton", undefined, "Scale by ratio conversion (e.g. 3/8\"=1'-0\" -> 1:10)");

    radioWidth.value = true; // Default selection

    function updateModeVisibility() {
        var isRatio = radioRatio.value;
        sizePanel.visible = !isRatio;
        ratioPanel.visible = isRatio;
        dialog.layout.layout(true);
        dialog.layout.resize();
        if (isRatio) updateRatioPreview();
    }

    function updateRatioPreview() {
        var sN = parseScaleInput(sourceInput.text);
        var tN = parseScaleInput(targetInput.text);
        if (sN && tN) {
            var factor = sN / tN;
            ratioPreview.text = "Source 1:" + formatNumber(sN) + "  ->  Target 1:" + formatNumber(tN) +
                                "    Scale factor: " + factor.toFixed(4) + "x  (" + (factor * 100).toFixed(2) + "%)";
        } else if (sourceInput.text || (targetInput.text && targetInput.text !== "1:10")) {
            ratioPreview.text = "(Enter both scales to see preview)";
        } else {
            ratioPreview.text = "";
        }
    }

    sourceInput.onChanging = updateRatioPreview;
    sourceInput.onChange = updateRatioPreview;
    targetInput.onChanging = updateRatioPreview;
    targetInput.onChange = updateRatioPreview;

    radioWidth.onClick = updateModeVisibility;
    radioHeight.onClick = updateModeVisibility;
    radioBoth.onClick = updateModeVisibility;
    radioFit.onClick = updateModeVisibility;
    radioRatio.onClick = updateModeVisibility;

    dialog.add("statictext", undefined, ""); // Spacer

    // Info text
    var infoText = dialog.add("statictext", undefined, "All selected objects will be scaled together.", {multiline: true});

    // Buttons
    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "center";
    buttonGroup.add("button", undefined, "OK", {name: "ok"});
    buttonGroup.add("button", undefined, "Cancel", {name: "cancel"});

    // Show dialog
    if (dialog.show() == 2) return; // User clicked Cancel

    // Calculate scale factors based on selected mode
    var scaleX, scaleY;
    var sourceScaleN = null, targetScaleN = null;

    if (radioRatio.value) {
        sourceScaleN = parseScaleInput(sourceInput.text);
        targetScaleN = parseScaleInput(targetInput.text);
        if (!sourceScaleN || sourceScaleN <= 0) {
            alert("Could not parse Source Scale.\n\nTry formats like:\n  3/8\" = 1'-0\"\n  1/4\" = 1'\n  1:32");
            return;
        }
        if (!targetScaleN || targetScaleN <= 0) {
            alert("Could not parse Target Scale.\n\nTry formats like:\n  1:10\n  1\" = 10\"\n  1\" = 1'");
            return;
        }
        scaleX = scaleY = (sourceScaleN / targetScaleN) * 100;
    } else {
        // Size-based modes
        var targetWidth = parseFloat(widthInput.text);
        var targetHeight = parseFloat(heightInput.text);

        if (radioWidth.value) {
            if (isNaN(targetWidth) || targetWidth <= 0) {
                alert("Please enter a valid positive number for Target Width.");
                return;
            }
        } else if (radioHeight.value) {
            if (isNaN(targetHeight) || targetHeight <= 0) {
                alert("Please enter a valid positive number for Target Height.");
                return;
            }
        } else {
            if (isNaN(targetWidth) || targetWidth <= 0 || isNaN(targetHeight) || targetHeight <= 0) {
                alert("Please enter valid positive numbers for both dimensions.");
                return;
            }
        }

        var targetWidthPts = targetWidth * 72;
        var targetHeightPts = targetHeight * 72;

        if (radioWidth.value) {
            scaleX = scaleY = (targetWidthPts / currentWidth) * 100;
        } else if (radioHeight.value) {
            scaleX = scaleY = (targetHeightPts / currentHeight) * 100;
        } else if (radioBoth.value) {
            scaleX = (targetWidthPts / currentWidth) * 100;
            scaleY = (targetHeightPts / currentHeight) * 100;
        } else if (radioFit.value) {
            var scaleByWidth = targetWidthPts / currentWidth;
            var scaleByHeight = targetHeightPts / currentHeight;
            scaleX = scaleY = Math.min(scaleByWidth, scaleByHeight) * 100;
        }
    }

    // Get the center point of the entire selection for scaling
    var selBounds = getSelectionBounds(sel);
    var centerX = (selBounds[0] + selBounds[2]) / 2;
    var centerY = (selBounds[1] + selBounds[3]) / 2;

    // Group all selected objects temporarily to maintain relative positioning
    var wasGrouped = false;
    var tempGroup;

    if (sel.length > 1) {
        // Check if selection is already a single group
        if (sel.length === 1 && sel[0].typename === "GroupItem") {
            tempGroup = sel[0];
            wasGrouped = true;
        } else {
            // Create temporary group
            tempGroup = doc.groupItems.add();
            for (var i = sel.length - 1; i >= 0; i--) {
                sel[i].moveToBeginning(tempGroup);
            }
        }

        // Scale the group as a whole
        tempGroup.resize(scaleX, scaleY, true, true, true, true, scaleX, Transformation.CENTER);

        // Ungroup if we created a temporary group
        if (!wasGrouped) {
            var groupItems = [];
            for (var i = 0; i < tempGroup.pageItems.length; i++) {
                groupItems.push(tempGroup.pageItems[i]);
            }
            for (var i = groupItems.length - 1; i >= 0; i--) {
                groupItems[i].moveToBeginning(tempGroup.parent);
            }
            tempGroup.remove();
        }
    } else {
        // Single object, scale normally
        sel[0].resize(scaleX, scaleY, true, true, true, true, scaleX, Transformation.CENTER);
    }

    // Report results
    var finalBounds = keyObject.geometricBounds;
    var finalWidth = (finalBounds[2] - finalBounds[0]) / 72;
    var finalHeight = (finalBounds[1] - finalBounds[3]) / 72;

    var report = "Scaling complete!\n\n";
    if (radioRatio.value) {
        report += "Ratio conversion:\n" +
                  "  Source: 1:" + formatNumber(sourceScaleN) + "\n" +
                  "  Target: 1:" + formatNumber(targetScaleN) + "\n" +
                  "  Scale factor: " + (scaleX / 100).toFixed(4) + "x\n\n";
    }
    report += "Key object scaled to:\n" +
              "Width: " + finalWidth.toFixed(4) + '"\n' +
              "Height: " + finalHeight.toFixed(4) + '"\n\n' +
              "Scale factors applied:\n" +
              "X: " + scaleX.toFixed(2) + "%\n" +
              "Y: " + scaleY.toFixed(2) + "%";

    alert(report);
}

// Helper function to get bounds of entire selection
function getSelectionBounds(selection) {
    var left = Infinity, top = -Infinity, right = -Infinity, bottom = Infinity;

    for (var i = 0; i < selection.length; i++) {
        var bounds = selection[i].geometricBounds;
        left = Math.min(left, bounds[0]);
        top = Math.max(top, bounds[1]);
        right = Math.max(right, bounds[2]);
        bottom = Math.min(bottom, bounds[3]);
    }

    return [left, top, right, bottom];
}

// ============================================================================
// SCALE INPUT PARSING
// ============================================================================
// Parses scale notation into a single number N representing ratio 1:N.
// Supported formats:
//   "1:32"           -> 32
//   "1 : 50"         -> 50
//   "3/8\" = 1'-0\"" -> 32     (3/8" page = 12" real, so 1:32)
//   "1/4\" = 1'"     -> 48
//   "1\" = 1'"       -> 12
//   "1\" = 10\""     -> 10
//   "1 1/2\" = 1'"   -> 8      (mixed numbers supported)
//   "0.375 = 12"     -> 32     (decimals, no inch marks ok too)
// Returns null if the input cannot be parsed.
function parseScaleInput(input) {
    if (input === undefined || input === null) return null;
    var s = String(input).replace(/^\s+|\s+$/g, '');
    if (s === '') return null;

    // Format: "1:N" or "1 : N"
    var ratioMatch = s.match(/^1\s*:\s*([\d.]+)\s*$/);
    if (ratioMatch) {
        var n = parseFloat(ratioMatch[1]);
        return (isNaN(n) || n <= 0) ? null : n;
    }

    // Architectural: "left = right"
    var eqIdx = s.indexOf('=');
    if (eqIdx > 0 && eqIdx < s.length - 1) {
        var leftStr = s.substring(0, eqIdx);
        var rightStr = s.substring(eqIdx + 1);
        var leftIn = parseAsInches(leftStr);
        var rightIn = parseAsInches(rightStr);
        if (leftIn === null || rightIn === null || leftIn <= 0 || rightIn <= 0) {
            return null;
        }
        return rightIn / leftIn;
    }

    return null;
}

// Parse an expression representing a length in inches.
// Handles: feet-inches ("1'-6\""), feet only ("1'"), fractions ("3/8"),
// mixed numbers ("1 1/2"), decimals ("0.375"), whole numbers ("12"),
// with or without quote marks.
function parseAsInches(text) {
    if (text === undefined || text === null) return null;
    var s = String(text).replace(/^\s+|\s+$/g, '');
    if (s === '') return null;

    var totalInches = 0;

    // Feet portion (e.g. "1'" or start of "1'-6\"")
    var feetMatch = s.match(/^(\d+(?:\.\d+)?)\s*'/);
    if (feetMatch) {
        totalInches += parseFloat(feetMatch[1]) * 12;
        s = s.substring(feetMatch[0].length);
        // Strip leading dash/whitespace before inches portion
        s = s.replace(/^[\s\-]+/, '');
    }

    // Strip trailing inch mark and whitespace
    s = s.replace(/"\s*$/, '').replace(/^\s+|\s+$/g, '');

    if (s === '') return totalInches;

    // Mixed number: "1 1/2"
    var mixed = s.match(/^(\d+)\s+(\d+)\s*\/\s*(\d+)$/);
    if (mixed) {
        var denom = parseFloat(mixed[3]);
        if (denom === 0) return null;
        totalInches += parseFloat(mixed[1]) + (parseFloat(mixed[2]) / denom);
        return totalInches;
    }

    // Simple fraction: "3/8"
    var frac = s.match(/^(\d+)\s*\/\s*(\d+)$/);
    if (frac) {
        var d = parseFloat(frac[2]);
        if (d === 0) return null;
        totalInches += parseFloat(frac[1]) / d;
        return totalInches;
    }

    // Decimal or whole number: "0.375", "12", "1.5"
    if (/^\d+(?:\.\d+)?$/.test(s)) {
        totalInches += parseFloat(s);
        return totalInches;
    }

    return null;
}

// Pretty-print a ratio denominator (avoids "32.0000000001" style noise)
function formatNumber(n) {
    if (n === Math.floor(n)) return n.toString();
    return n.toFixed(4).replace(/0+$/, '').replace(/\.$/, '');
}

// Run the script
main();
