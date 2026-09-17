/*@METADATA{
  "name": "Image Quality Analysis",
  "description": "Analyze selected images and label them with resolution info; auto-detects the drawing scale from the page",
  "version": "2.2",
  "target": "illustrator",
  "tags": ["image", "quality", "resolution", "PPI"]
}@END_METADATA*/

function main() {
    if (app.documents.length === 0) {
        alert("Please open a document first.");
        return;
    }
    
    var doc = app.activeDocument;
    var sel = doc.selection;
    
    if (sel.length === 0) {
        alert("Please select one or more objects containing images.");
        return;
    }
    
    // Look for a scale callout on the page so the dialog can prefill it
    var detectedScale = detectPageScale(doc);

    // Show scale selection dialog
    var scaleDialog = new Window("dialog", "Image Quality Analysis - Select Scale");
    scaleDialog.alignChildren = "fill";
    scaleDialog.spacing = 15;
    scaleDialog.margins = 20;

    var titleText = scaleDialog.add("statictext", undefined, "Select the drawing scale:");
    titleText.graphics.font = ScriptUI.newFont("dialog", "Bold", 12);

    var scaleItems = ["1:1", "1:2", "1:4", "1:5", "1:8", "1:10", "1:16", "1:20", "1:40", "1:50", "1:100", "2:1", "4:1", "8:1", "10:1", "100:2", "Custom"];
    var customIndex = scaleItems.length - 1;

    var scaleGroup = scaleDialog.add("group");
    scaleGroup.add("statictext", undefined, "Scale:");
    var scaleDropdown = scaleGroup.add("dropdownlist", undefined, scaleItems);
    scaleDropdown.selection = 5; // Default to 1:10
    scaleDropdown.preferredSize.width = 100;

    var customGroup = scaleGroup.add("group");
    var customScaleNumerator = customGroup.add("edittext", undefined, "1");
    customScaleNumerator.characters = 3;
    customScaleNumerator.enabled = false;
    
    customGroup.add("statictext", undefined, ":");
    
    var customScaleDenominator = customGroup.add("edittext", undefined, "1");
    customScaleDenominator.characters = 3;
    customScaleDenominator.enabled = false;
    
    scaleDropdown.onChange = function() {
        var isCustom = scaleDropdown.selection.text === "Custom";
        customScaleNumerator.enabled = isCustom;
        customScaleDenominator.enabled = isCustom;
    };

    // Status line: what was (or was not) found on the page
    var statusText = scaleDialog.add("statictext", undefined, "", {multiline: true});
    statusText.preferredSize.width = 360;
    statusText.preferredSize.height = 30;

    if (detectedScale) {
        var matchIndex = -1;
        for (var si = 0; si < scaleItems.length; si++) {
            if (scaleItems[si] === detectedScale.text) {
                matchIndex = si;
                break;
            }
        }

        if (matchIndex >= 0) {
            scaleDropdown.selection = matchIndex;
        } else {
            // Not in the preset list - drop it into the Custom fields
            scaleDropdown.selection = customIndex;
            var detParts = detectedScale.text.split(":");
            customScaleNumerator.text = detParts[0];
            customScaleDenominator.text = detParts[1];
            customScaleNumerator.enabled = true;
            customScaleDenominator.enabled = true;
        }

        statusText.text = "Found on page: " + detectedScale.text +
            '   (from "' + truncateForDisplay(detectedScale.source, 40) + '")';
    } else {
        statusText.text = "No scale found on the page - using the default 1:10.";
    }

    var buttonGroup = scaleDialog.add("group");
    buttonGroup.alignment = "center";
    buttonGroup.add("button", undefined, "OK", {name: "ok"});
    buttonGroup.add("button", undefined, "Cancel", {name: "cancel"});
    
    if (scaleDialog.show() == 2) return;
    
    // Get scale settings
    var scaleText;
    var scaleRatio;
    if (scaleDropdown.selection.text === "Custom") {
        scaleText = customScaleNumerator.text + ":" + customScaleDenominator.text;
        var num = parseFloat(customScaleNumerator.text) || 1;
        var denom = parseFloat(customScaleDenominator.text) || 1;
        scaleRatio = denom / num;
    } else {
        scaleText = scaleDropdown.selection.text;
        var parts = scaleText.split(":");
        scaleRatio = parseFloat(parts[1]) / parseFloat(parts[0]);
    }
    
    // Get document scale factor for Large Canvas support
    var scaleFactor = 1;
    try {
        scaleFactor = doc.scaleFactor || 1;
    } catch (e) {
        scaleFactor = 1;
    }
    
    // Create or get Image Analysis layer
    var analysisLayer = getOrCreateLayer("Image Analysis");
    
    // Collect all clipping masks from selection
    var clippingMasks = [];
    collectClippingMasks(sel, clippingMasks);
    
    // Process each item (clipping mask or standalone image)
    var labelsCreated = 0;
    for (var i = 0; i < clippingMasks.length; i++) {
        var item = clippingMasks[i];
        
        // Get the VISIBLE bounds using the artboard method
        var visibleBounds = getVisibleBounds(item);
        
        var result = processClippingMask(item, visibleBounds, analysisLayer, scaleRatio, scaleFactor);
        if (result) labelsCreated++;
    }
    
    if (labelsCreated === 0) {
        alert("No images found in selection.");
    } else {
        alert("Analysis complete!\n" + labelsCreated + " label(s) created.");
    }
}

// Get visible bounds using temporary artboard and Fit to Selected Art
function getVisibleBounds(obj) {
    var doc = app.activeDocument;
    
    try {
        // Store the current artboard index
        var originalArtboardIndex = doc.artboards.getActiveArtboardIndex();
        
        // Create a temporary artboard
        var tempArtboard = doc.artboards.add([0, 0, 100, -100]);
        var tempArtboardIndex = doc.artboards.length - 1;
        
        // Set the temp artboard as active
        doc.artboards.setActiveArtboardIndex(tempArtboardIndex);
        
        // Select only the current object
        doc.selection = null;
        obj.selected = true;
        
        // Use Fit Artboard to Selected Art menu command
        app.executeMenuCommand("Fit Artboard to selected Art");
        
        // Refresh to ensure artboard has updated
        app.redraw();
        
        // Get the artboard rect AFTER the fit command
        var artboardRect = doc.artboards[tempArtboardIndex].artboardRect;
        
        // Convert artboard rect to bounds format [left, top, right, bottom]
        var bounds = [
            artboardRect[0], // left
            artboardRect[1], // top
            artboardRect[2], // right
            artboardRect[3]  // bottom
        ];
        
        // Remove the temporary artboard
        doc.artboards.remove(tempArtboardIndex);
        
        // Restore the original active artboard
        if (originalArtboardIndex >= 0 && originalArtboardIndex < doc.artboards.length) {
            doc.artboards.setActiveArtboardIndex(originalArtboardIndex);
        }
        
        // Deselect the object
        obj.selected = false;
        
        return bounds;
        
    } catch (e) {
        // If anything fails, clean up and fall back to geometric bounds
        try {
            if (doc.artboards.length > 0) {
                var lastIndex = doc.artboards.length - 1;
                doc.artboards.remove(lastIndex);
            }
        } catch (cleanupError) {}
        
        try {
            if (originalArtboardIndex >= 0 && originalArtboardIndex < doc.artboards.length) {
                doc.artboards.setActiveArtboardIndex(originalArtboardIndex);
            }
        } catch (cleanupError) {}
        
        return obj.geometricBounds;
    }
}

function getOrCreateLayer(layerName) {
    var doc = app.activeDocument;
    try {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === layerName) {
                return doc.layers[i];
            }
        }
    } catch (e) {}
    
    var newLayer = doc.layers.add();
    newLayer.name = layerName;
    return newLayer;
}

function collectClippingMasks(items, results) {
    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        
        // Skip groups named "MAA Symbol ICI"
        if (item.typename === "GroupItem" && item.name === "MAA Symbol ICI") {
            continue;
        }
        
        // Add clipping masks
        if (item.typename === "GroupItem" && item.clipped) {
            results.push(item);
        }
        // Add standalone images
        else if (item.typename === "RasterItem" || item.typename === "PlacedItem") {
            results.push(item);
        }
        
        // Check inside groups for nested items
        if (item.typename === "GroupItem" && item.pageItems && item.pageItems.length > 0) {
            collectClippingMasks(item.pageItems, results);
        }
    }
}

function processClippingMask(clippedGroup, visibleBounds, layer, scaleRatio, scaleFactor) {
    try {
        // visibleBounds is already passed in - calculated using artboard method
        
        // Find the first image inside this clipped group
        var imageItem = findFirstImage(clippedGroup);
        
        if (!imageItem) {
            return false;
        }
        
        var actualScaleFactor = (scaleFactor && scaleFactor > 0) ? scaleFactor : 1;
        
        // Calculate PPI using matrix method
        var estimatedPPI = "Unknown";
        var isLowRes = false;
        
        try {
            var matrix = imageItem.matrix;
            if (matrix && matrix.mValueA !== undefined && matrix.mValueD !== undefined) {
                var scaleX = Math.abs(matrix.mValueA);
                var scaleY = Math.abs(matrix.mValueD);
                var avgScale = (scaleX + scaleY) / 2;
                
                if (avgScale > 0) {
                    var basePPI = 72 / avgScale;
                    var correctedPPI = basePPI / actualScaleFactor;
                    correctedPPI = correctedPPI / scaleRatio;
                    estimatedPPI = Math.round(correctedPPI);
                    isLowRes = (estimatedPPI < 72);
                }
            }
        } catch (matrixError) {}
        
        if (estimatedPPI !== "Unknown") {
            // Use the visibleBounds that was passed in
            createImageLabel(visibleBounds, estimatedPPI, isLowRes, layer);
            return true;
        }
        
    } catch (e) {
        // Silent failure for individual items
    }
    
    return false;
}

function findFirstImage(item) {
    if (item.typename === "RasterItem" || item.typename === "PlacedItem") {
        return item;
    }
    
    if (item.typename === "GroupItem" && item.pageItems && item.pageItems.length > 0) {
        for (var i = 0; i < item.pageItems.length; i++) {
            var found = findFirstImage(item.pageItems[i]);
            if (found) return found;
        }
    }
    
    return null;
}

function createImageLabel(bounds, ppi, isLowRes, layer) {
    // Bounds format: [left, top, right, bottom]
    var centerX = (bounds[0] + bounds[2]) / 2;
    var centerY = (bounds[1] + bounds[3]) / 2;
    
    // Create text frame
    var label = layer.textFrames.add();
    
    // Simple format: just XXXPPI
    label.contents = ppi + 'PPI';
    
    // Set font to Myriad Pro Black 30pt
    try {
        label.textRange.characterAttributes.textFont = app.textFonts.getByName("MyriadPro-Black");
    } catch (e) {
        try {
            label.textRange.characterAttributes.textFont = app.textFonts.getByName("Myriad-Black");
        } catch (e2) {
            try {
                label.textRange.characterAttributes.textFont = app.textFonts.getByName("MyriadPro-Bold");
            } catch (e3) {}
        }
    }
    
    label.textRange.characterAttributes.size = 30;
    
    // Set center alignment
    label.textRange.paragraphAttributes.justification = Justification.CENTER;
    
    // Set colors based on quality
    if (isLowRes) {
        // Red fill with black stroke for low resolution
        var redColor = new CMYKColor();
        redColor.cyan = 0;
        redColor.magenta = 100;
        redColor.yellow = 100;
        redColor.black = 0;
        
        var blackColor = new CMYKColor();
        blackColor.cyan = 0;
        blackColor.magenta = 0;
        blackColor.yellow = 0;
        blackColor.black = 100;
        
        label.textRange.characterAttributes.strokeColor = blackColor;
        label.textRange.characterAttributes.stroked = true;
        label.textRange.characterAttributes.lineWidth = 3;
        
        label.textRange.characterAttributes.fillColor = redColor;
        label.textRange.characterAttributes.filled = true;
    } else {
        // Black fill with white stroke for good resolution
        var blackColor = new CMYKColor();
        blackColor.cyan = 0;
        blackColor.magenta = 0;
        blackColor.yellow = 0;
        blackColor.black = 100;
        
        var whiteColor = new CMYKColor();
        whiteColor.cyan = 0;
        whiteColor.magenta = 0;
        whiteColor.yellow = 0;
        whiteColor.black = 0;
        
        label.textRange.characterAttributes.strokeColor = whiteColor;
        label.textRange.characterAttributes.stroked = true;
        label.textRange.characterAttributes.lineWidth = 3;
        
        label.textRange.characterAttributes.fillColor = blackColor;
        label.textRange.characterAttributes.filled = true;
    }
    
    // Position text: center it on the bounds
    label.left = centerX - (label.width / 2);
    label.top = centerY + (label.height / 2);
}

// ============================================================================
// SCALE DETECTION
// ============================================================================
// Scan the document's text for a drawing scale callout so the dialog can
// prefill it. Recognizes ratio notation ("Scale 1:10", "1:20") and
// architectural notation ('3/8" = 1'-0"', '1/4" = 1\''). Returns
// {text: "1:10", ratio: 10, source: "Scale 1:10"} or null if nothing parses.
function detectPageScale(doc) {
    var frames;
    try {
        frames = doc.textFrames;
    } catch (e) {
        return null;
    }
    if (!frames || frames.length === 0) return null;

    // Bounds of the active artboard, used to prefer callouts on the current page
    var artRect = null;
    try {
        var abIndex = doc.artboards.getActiveArtboardIndex();
        if (abIndex >= 0 && abIndex < doc.artboards.length) {
            artRect = doc.artboards[abIndex].artboardRect;
        }
    } catch (e) {
        artRect = null;
    }

    var limit = frames.length;
    if (limit > 3000) limit = 3000; // sanity cap on very heavy documents

    var best = null;
    var bestScore = -1;

    for (var i = 0; i < limit; i++) {
        var contents;
        try {
            contents = frames[i].contents;
        } catch (e) {
            continue;
        }
        if (!contents) continue;

        var onArtboard = false;
        if (artRect) {
            try {
                var b = frames[i].geometricBounds; // [left, top, right, bottom]
                var cx = (b[0] + b[2]) / 2;
                var cy = (b[1] + b[3]) / 2;
                onArtboard = (cx >= artRect[0] && cx <= artRect[2] &&
                              cy <= artRect[1] && cy >= artRect[3]);
            } catch (e) {}
        }

        var lines = String(contents).split(/[\r\n]+/);
        for (var j = 0; j < lines.length; j++) {
            var hit = scaleFromLine(lines[j]);
            if (!hit) continue;

            // A line that actually says "Scale" beats a bare ratio, and a
            // callout on the active artboard beats one somewhere else.
            var score = (hit.labeled ? 2 : 0) + (onArtboard ? 1 : 0);
            if (score > bestScore) {
                bestScore = score;
                best = hit;
            }
        }
    }

    return best;
}

// Pull a scale out of a single line of text. Returns
// {text, ratio, labeled, source} or null.
function scaleFromLine(line) {
    var raw = normalizeScaleText(line).replace(/^\s+|\s+$/g, '');
    if (raw === '') return null;

    var s = raw;
    var labeled = false;

    // "Scale: 1:10", "SCALE 1/4" = 1'-0"", "Drawn to scale - 1:20"
    var labelMatch = s.match(/scale[\s:=\-]+(.+)$/i);
    if (labelMatch) {
        labeled = true;
        s = labelMatch[1];
    }

    // Trim surrounding punctuation and whitespace
    s = s.replace(/^[\s\(\[]+/, '').replace(/[\s\)\]\.,;]+$/, '');
    if (s === '') return null;

    var parsed = parseScaleExpression(s);
    if (!parsed) return null;

    parsed.labeled = labeled;
    parsed.source = raw;
    return parsed;
}

// Parse a scale expression into {text: "A:B", ratio: N} where N is the number
// the on-page measurement is multiplied by to get real-world size.
//   "1:10"            -> ratio 10
//   "2:1"             -> ratio 0.5
//   "3/8\" = 1'-0\""  -> ratio 32  (text "1:32")
//   "1\" = 1'"        -> ratio 12  (text "1:12")
function parseScaleExpression(text) {
    var s = normalizeScaleText(text).replace(/^\s+|\s+$/g, '');
    if (s === '') return null;

    // Ratio form: "A:B"
    var ratioMatch = s.match(/^(\d+(?:\.\d+)?)\s*:\s*(\d+(?:\.\d+)?)$/);
    if (ratioMatch) {
        var a = parseFloat(ratioMatch[1]);
        var b = parseFloat(ratioMatch[2]);
        if (isNaN(a) || isNaN(b) || a <= 0 || b <= 0) return null;
        return { text: formatNumber(a) + ":" + formatNumber(b), ratio: b / a };
    }

    // Architectural form: "left = right"
    var eqIdx = s.indexOf('=');
    if (eqIdx > 0 && eqIdx < s.length - 1) {
        var leftIn = parseAsInches(s.substring(0, eqIdx));
        var rightIn = parseAsInches(s.substring(eqIdx + 1));
        if (leftIn === null || rightIn === null || leftIn <= 0 || rightIn <= 0) {
            return null;
        }
        var r = rightIn / leftIn;
        return { text: "1:" + formatNumber(r), ratio: r };
    }

    return null;
}

// Parse an expression representing a length in inches.
// Handles feet-inches ("1'-6\""), feet only ("1'"), fractions ("3/8"),
// mixed numbers ("1 1/2" or "1-1/2"), decimals, and whole numbers.
function parseAsInches(text) {
    if (text === undefined || text === null) return null;
    var s = normalizeScaleText(text).replace(/^\s+|\s+$/g, '');
    if (s === '') return null;

    var totalInches = 0;

    // Feet portion (e.g. "1'" or the start of "1'-6\"")
    var feetMatch = s.match(/^(\d+(?:\.\d+)?)\s*'/);
    if (feetMatch) {
        totalInches += parseFloat(feetMatch[1]) * 12;
        s = s.substring(feetMatch[0].length);
        s = s.replace(/^[\s\-]+/, '');
    }

    // Strip trailing inch mark and whitespace
    s = s.replace(/"\s*$/, '').replace(/^\s+|\s+$/g, '');

    if (s === '') return totalInches;

    // Mixed number: "1 1/2" or "1-1/2"
    var mixed = s.match(/^(\d+)[\s\-]+(\d+)\s*\/\s*(\d+)$/);
    if (mixed) {
        var mDenom = parseFloat(mixed[3]);
        if (mDenom === 0) return null;
        totalInches += parseFloat(mixed[1]) + (parseFloat(mixed[2]) / mDenom);
        return totalInches;
    }

    // Simple fraction: "3/8"
    var frac = s.match(/^(\d+)\s*\/\s*(\d+)$/);
    if (frac) {
        var fDenom = parseFloat(frac[2]);
        if (fDenom === 0) return null;
        totalInches += parseFloat(frac[1]) / fDenom;
        return totalInches;
    }

    // Decimal or whole number: "0.375", "12", "1.5"
    if (/^\d+(?:\.\d+)?$/.test(s)) {
        totalInches += parseFloat(s);
        return totalInches;
    }

    return null;
}

// Normalize the typographic characters that show up in real proofs (prime and
// double-prime marks, curly quotes, en/em dashes, non-breaking spaces, Unicode
// fraction glyphs) down to plain ASCII. Done by char code so this source file
// stays pure ASCII.
function normalizeScaleText(input) {
    if (input === undefined || input === null) return '';
    var src = String(input);
    var out = [];
    for (var i = 0; i < src.length; i++) {
        var c = src.charCodeAt(i);
        if (c === 0x201C || c === 0x201D || c === 0x201E || c === 0x201F ||
            c === 0x2033 || c === 0x3003 || c === 0x301D || c === 0x301E || c === 0x301F) {
            out.push('"');   // double-prime / smart double quotes -> "
        } else if (c === 0x2018 || c === 0x2019 || c === 0x201A || c === 0x201B ||
                   c === 0x2032 || c === 0x00B4 || c === 0x0060) {
            out.push("'");   // prime / smart single quotes / acute / backtick -> '
        } else if (c === 0x2010 || c === 0x2011 || c === 0x2012 || c === 0x2013 ||
                   c === 0x2014 || c === 0x2015 || c === 0x2212) {
            out.push('-');   // assorted dashes / minus -> hyphen
        } else if (c === 0x00A0 || (c >= 0x2000 && c <= 0x200B) ||
                   c === 0x202F || c === 0x205F || c === 0x3000) {
            out.push(' ');   // assorted spaces -> space
        } else {
            var fr = fractionForCode(c);
            out.push(fr !== null ? (' ' + fr + ' ') : src.charAt(i));
        }
    }
    return out.join('');
}

function fractionForCode(c) {
    switch (c) {
        case 0x00BC: return '1/4';
        case 0x00BD: return '1/2';
        case 0x00BE: return '3/4';
        case 0x2153: return '1/3';
        case 0x2154: return '2/3';
        case 0x2155: return '1/5';
        case 0x2156: return '2/5';
        case 0x2157: return '3/5';
        case 0x2158: return '4/5';
        case 0x2159: return '1/6';
        case 0x215A: return '5/6';
        case 0x215B: return '1/8';
        case 0x215C: return '3/8';
        case 0x215D: return '5/8';
        case 0x215E: return '7/8';
    }
    return null;
}

// Pretty-print a ratio value (avoids "32.0000000001" style noise)
function formatNumber(n) {
    if (n === Math.floor(n)) return n.toString();
    return n.toFixed(4).replace(/0+$/, '').replace(/\.$/, '');
}

function truncateForDisplay(text, maxLen) {
    var s = String(text || '').replace(/^\s+|\s+$/g, '');
    if (s.length <= maxLen) return s;
    return s.substring(0, maxLen - 3) + '...';
}

// Run the script
main();