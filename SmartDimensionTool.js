/*@METADATA{
  "name": "Smart Dimension Tool",
  "description": "Add dimensions to an array of signs without the fuss",
  "version": "5.0",
  "target": "illustrator",
  "tags": ["Measure", "Smart", "Utility"]
}@END_METADATA*/

// ===== Measurement formatting helpers =====

// Fraction slash glyph (U+2044) -- kerns tightly against numerator/denominator
var FRACTION_SLASH = String.fromCharCode(8260);

// Greatest common divisor (for reducing fractions)
function gcd(a, b) {
    a = Math.abs(a);
    b = Math.abs(b);
    while (b > 0) {
        var t = b;
        b = a % b;
        a = t;
    }
    return a;
}

// Convert a decimal inch value to a fraction string like: 12 3/8  (no quote mark)
// The value is rounded to the nearest 1/denominator first, then reduced.
function toFractionString(value, denominator) {
    if (!denominator || denominator < 1) denominator = 16;
    var negative = value < 0;
    var v = Math.abs(value);
    var totalUnits = Math.round(v * denominator);
    var whole = Math.floor(totalUnits / denominator);
    var numer = totalUnits - (whole * denominator);
    var result;
    if (numer === 0) {
        result = String(whole);
    } else {
        var g = gcd(numer, denominator);
        var fracStr = (numer / g) + "/" + (denominator / g);
        result = (whole > 0) ? (whole + " " + fracStr) : fracStr;
    }
    return negative ? "-" + result : result;
}

// Format a decimal value to a fixed precision, then strip trailing zeros
function toDecimalString(value, precision) {
    var s = value.toFixed(precision);
    if (s.indexOf('.') > -1) {
        s = s.replace(/\.?0+$/, '');
    }
    return s;
}

// Format a dimension value (inches) per the current settings:
// applies rounding, then renders as a fraction or decimal.
// Fractions require a rounding increment -- if rounding is "None",
// the value always renders as decimal.
function formatDimensionValue(value, settings) {
    var rounded = value;
    if (settings.roundingIncrement > 0) {
        rounded = Math.round(value / settings.roundingIncrement) * settings.roundingIncrement;
    }
    if (settings.useFractions && settings.roundingIncrement > 0) {
        var denom = Math.round(1 / settings.roundingIncrement);
        if (denom < 1) denom = 1;
        return toFractionString(rounded, denom);
    }
    return toDecimalString(rounded, settings.precision);
}

// ===== Typographic fraction styling =====

// Get the full-size point size of a measurement label. Uses the last
// character (the inch mark, which is never shrunk) so it works on frames
// that already have fraction styling applied.
function getBaseTextSize(textFrame) {
    try {
        var chars = textFrame.characters;
        return chars[chars.length - 1].characterAttributes.size;
    } catch (e) {}
    try {
        return textFrame.textRange.characterAttributes.size;
    } catch (e2) {}
    return 12;
}

// Reset a text frame to uniform styling (used before re-writing contents,
// so leftover small/raised fraction characters don't bleed into new text)
function resetTextStyling(textFrame, baseSize) {
    try {
        textFrame.textRange.characterAttributes.size = baseSize;
        textFrame.textRange.characterAttributes.baselineShift = 0;
    } catch (e) {}
}

// Turn the "N/D" part of a measurement label into a typographic fraction:
// numerator shrunk and raised, slash swapped for the fraction-slash glyph,
// denominator shrunk on the baseline. The space between the whole number
// and the fraction is also shrunk so the label reads as one unit.
// Safe to call on labels with no fraction (does nothing).
function styleFractionText(textFrame, baseSize) {
    try {
        var content = String(textFrame.contents);
        var normalized = content.split(FRACTION_SLASH).join("/");
        var m = normalized.match(/^(?:(\d+)[ -])?(\d+)\/(\d+)\s*"$/);
        if (!m) return;

        var wholeLen = m[1] ? m[1].length + 1 : 0; // +1 for the separator
        var numLen = m[2].length;
        var denLen = m[3].length;
        var slashIdx = wholeLen + numLen;

        // Swap a plain slash for the fraction-slash glyph
        if (content.charAt(slashIdx) !== FRACTION_SLASH) {
            textFrame.characters[slashIdx].contents = FRACTION_SLASH;
        }

        var smallSize = baseSize * 0.62;
        var i;

        // Shrink the separator space so the fraction tucks in close
        if (wholeLen > 0) {
            textFrame.characters[wholeLen - 1].characterAttributes.size = smallSize;
        }
        // Numerator: small and raised
        for (i = wholeLen; i < wholeLen + numLen; i++) {
            textFrame.characters[i].characterAttributes.size = smallSize;
            textFrame.characters[i].characterAttributes.baselineShift = baseSize * 0.31;
        }
        // Denominator: small, on the baseline
        for (i = slashIdx + 1; i < slashIdx + 1 + denLen; i++) {
            textFrame.characters[i].characterAttributes.size = smallSize;
            textFrame.characters[i].characterAttributes.baselineShift = 0;
        }
    } catch (e) {}
}

// ===== Existing-dimension conversion (runs when nothing is selected) =====

// Try to parse a text frame's contents as a measurement string.
// Returns { type: "decimal"|"fraction", value: <inches> } or null.
// Matches:  12.375"   12"   3/8"   12 3/8"   12-3/8"
function parseMeasurementText(content) {
    var s = String(content).replace(/^\s+|\s+$/g, '');
    // Normalize the fraction-slash glyph so styled fractions parse too
    s = s.split(FRACTION_SLASH).join("/");
    // Decimal (or whole number) like 12.375" or 12"
    var mDec = s.match(/^(\d+(?:\.\d+)?)\s*"$/);
    if (mDec) {
        return { type: "decimal", value: parseFloat(mDec[1]), isWhole: s.indexOf('.') === -1 && s.indexOf('/') === -1 };
    }
    // Fraction like 12 3/8", 12-3/8", or 3/8"
    var mFrac = s.match(/^(?:(\d+)[ -])?(\d+)\/(\d+)\s*"$/);
    if (mFrac) {
        var whole = mFrac[1] ? parseInt(mFrac[1], 10) : 0;
        var n = parseInt(mFrac[2], 10);
        var d = parseInt(mFrac[3], 10);
        if (d === 0) return null;
        return { type: "fraction", value: whole + (n / d) };
    }
    return null;
}

// Scan the document for measurement text frames on the dimensions layer
// and convert them between decimal and fraction formats.
function convertDimensionTexts() {
    var doc = app.activeDocument;
    var LAYER_NAME = "Dimensions";

    // Find the dimensions layer
    var dimLayer = null;
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === LAYER_NAME) {
            dimLayer = doc.layers[i];
            break;
        }
    }
    if (!dimLayer) {
        alert("Nothing is selected and no \"" + LAYER_NAME + "\" layer was found.\n\n" +
              "Select objects to add dimensions, or open a document that has a \"" +
              LAYER_NAME + "\" layer to convert existing measurements.");
        return;
    }

    // Collect measurement text frames on that layer (includes frames nested
    // inside dimension groups -- doc.textFrames covers the whole document)
    var found = [];
    var decimalCount = 0;
    var fractionCount = 0;
    for (var i = 0; i < doc.textFrames.length; i++) {
        var tf = doc.textFrames[i];
        var onLayer = false;
        try {
            onLayer = (tf.layer.name === LAYER_NAME);
        } catch (e) {}
        if (!onLayer) continue;

        var parsed = parseMeasurementText(tf.contents);
        if (!parsed) continue;

        found.push({ frame: tf, parsed: parsed });
        if (parsed.type === "decimal" && !parsed.isWhole) {
            decimalCount++;
        } else if (parsed.type === "fraction") {
            fractionCount++;
        }
    }

    if (found.length === 0) {
        alert("No measurement text (like 12.375\" or 12 3/8\") was found on the \"" +
              LAYER_NAME + "\" layer.");
        return;
    }

    // Build the conversion dialog
    var dlg = new Window("dialog", "Convert Dimension Text");
    dlg.alignChildren = "fill";

    var infoPanel = dlg.add("panel", undefined, "Found on \"" + LAYER_NAME + "\" layer");
    infoPanel.alignChildren = "left";
    infoPanel.add("statictext", undefined,
        found.length + " measurement label(s): " + decimalCount + " decimal, " +
        fractionCount + " fraction" + (found.length - decimalCount - fractionCount > 0 ?
        ", " + (found.length - decimalCount - fractionCount) + " whole number" : ""));

    var dirPanel = dlg.add("panel", undefined, "Convert");
    dirPanel.alignChildren = "left";
    var toFractionsRadio = dirPanel.add("radiobutton", undefined, "Decimal -> Fractions  (12.375\" -> 12 3/8\")");
    var toDecimalsRadio = dirPanel.add("radiobutton", undefined, "Fractions -> Decimal  (12 3/8\" -> 12.375\")");

    // Smart default: convert whichever style is more common into the other
    if (fractionCount > decimalCount) {
        toDecimalsRadio.value = true;
    } else {
        toFractionsRadio.value = true;
    }

    var precRow = dirPanel.add("group");
    precRow.add("statictext", undefined, "Fraction precision:");
    var precDropdown = precRow.add("dropdownlist", undefined, ["1/2", "1/4", "1/8", "1/16", "1/32"]);
    precDropdown.selection = 3; // Default to 1/16
    precDropdown.enabled = toFractionsRadio.value;

    toFractionsRadio.onClick = function() { precDropdown.enabled = true; };
    toDecimalsRadio.onClick = function() { precDropdown.enabled = false; };

    var btns = dlg.add("group");
    btns.add("button", undefined, "OK", { name: "ok" });
    btns.add("button", undefined, "Cancel", { name: "cancel" });

    if (dlg.show() == 2) return;

    var convertedCount = 0;
    var restyledCount = 0;

    if (toFractionsRadio.value) {
        // Decimal -> Fractions. Also re-renders existing fraction labels so
        // plain ones (e.g. from older versions) pick up the typographic style.
        var denom = parseInt(precDropdown.selection.text.split("/")[1], 10) || 16;
        for (var i = 0; i < found.length; i++) {
            var entry = found[i];
            if (entry.parsed.type === "decimal" && entry.parsed.isWhole) continue;

            var baseSize = getBaseTextSize(entry.frame);
            var newPlain = toFractionString(entry.parsed.value, denom) + '"';
            var oldNormalized = String(entry.frame.contents).split(FRACTION_SLASH).join("/");
            var alreadyStyled = String(entry.frame.contents).indexOf(FRACTION_SLASH) > -1;

            if (oldNormalized !== newPlain || !alreadyStyled) {
                resetTextStyling(entry.frame, baseSize);
                entry.frame.contents = newPlain;
                styleFractionText(entry.frame, baseSize);
                if (entry.parsed.type === "decimal" || oldNormalized !== newPlain) {
                    convertedCount++;
                } else {
                    restyledCount++;
                }
            }
        }
    } else {
        // Fractions -> Decimal (5 decimals covers 1/32, trailing zeros stripped)
        for (var i = 0; i < found.length; i++) {
            var entry = found[i];
            if (entry.parsed.type !== "fraction") continue;

            var baseSize = getBaseTextSize(entry.frame);
            var newText = toDecimalString(entry.parsed.value, 5) + '"';
            if (newText !== entry.frame.contents) {
                entry.frame.contents = newText;
                resetTextStyling(entry.frame, baseSize);
                convertedCount++;
            }
        }
    }

    app.redraw();
    var report = "Converted " + convertedCount + " measurement label(s).";
    if (restyledCount > 0) {
        report += "\nRestyled " + restyledCount + " existing fraction label(s).";
    }
    alert(report);
}

// Get character style by name
function getCharacterStyle(styleName) {
    var doc = app.activeDocument;
    try {
        for (var i = 0; i < doc.characterStyles.length; i++) {
            if (doc.characterStyles[i].name === styleName) {
                return doc.characterStyles[i];
            }
        }
    } catch (e) {}
    return null;
}

// Detect existing scale text on artboards containing selected objects
function detectExistingScale(selection) {
    var doc = app.activeDocument;
    var artboardIndices = [];
    
    // Find which artboards contain the selected objects
    for (var i = 0; i < selection.length; i++) {
        var itemBounds = selection[i].geometricBounds;
        var itemCenterX = (itemBounds[0] + itemBounds[2]) / 2;
        var itemCenterY = (itemBounds[1] + itemBounds[3]) / 2;
        
        for (var j = 0; j < doc.artboards.length; j++) {
            var artboardRect = doc.artboards[j].artboardRect;
            if (itemCenterX >= artboardRect[0] && itemCenterX <= artboardRect[2] &&
                itemCenterY <= artboardRect[1] && itemCenterY >= artboardRect[3]) {
                var alreadyAdded = false;
                for (var k = 0; k < artboardIndices.length; k++) {
                    if (artboardIndices[k] === j) {
                        alreadyAdded = true;
                        break;
                    }
                }
                if (!alreadyAdded) {
                    artboardIndices.push(j);
                }
                break;
            }
        }
    }
    
    // Search for scale text on these artboards
    for (var i = 0; i < artboardIndices.length; i++) {
        var artboard = doc.artboards[artboardIndices[i]];
        var artboardRect = artboard.artboardRect;
        
        // Check all text frames in the document
        for (var j = 0; j < doc.textFrames.length; j++) {
            var textFrame = doc.textFrames[j];
            var textBounds = textFrame.geometricBounds;
            var textCenterX = (textBounds[0] + textBounds[2]) / 2;
            var textCenterY = (textBounds[1] + textBounds[3]) / 2;
            
            // Check if text is on this artboard
            if (textCenterX >= artboardRect[0] && textCenterX <= artboardRect[2] &&
                textCenterY <= artboardRect[1] && textCenterY >= artboardRect[3]) {
                
                // Check if text starts with "Scale " (case insensitive)
                var content = textFrame.contents;
                if (content.toLowerCase().indexOf("scale ") === 0) {
                    // Extract the scale value (everything after "Scale ")
                    var scaleValue = content.substring(6);
                    // Manual trim - remove leading/trailing whitespace
                    scaleValue = scaleValue.replace(/^\s+|\s+$/g, '');
                    if (scaleValue.length > 0) {
                        return scaleValue;
                    }
                }
            }
        }
    }
    
    return null;
}

// Add below-art texts: a per-sign "Qty: XX" label below each selected item
// and/or one "Below Centered" Scale label per artboard group. Both go on
// the Dimensions layer (dimLayer), alongside the dimension markers.
//
// Spacing matches the dimension marker spacing -- settings.offset (0.1") gap
// between the art (or bottom-dim text) and the Qty label, and the same gap
// between the lowest Qty (or art/dim) and the Scale label.
function addBelowArtTexts(selection, settings, dimLayer) {
    var doc = app.activeDocument;

    // Group selected items by the artboard that contains their center, so
    // Scale "Below Centered" gets one label per artboard centered on its art.
    // Items not on any artboard share a single "no artboard" bucket.
    var artboardGroups = {};
    var artboardOrder = [];
    var noArtboardItems = [];

    for (var i = 0; i < selection.length; i++) {
        var itemBounds = selection[i].geometricBounds;
        var itemCenterX = (itemBounds[0] + itemBounds[2]) / 2;
        var itemCenterY = (itemBounds[1] + itemBounds[3]) / 2;

        var assigned = false;
        for (var j = 0; j < doc.artboards.length; j++) {
            var artboardRect = doc.artboards[j].artboardRect;
            if (itemCenterX >= artboardRect[0] && itemCenterX <= artboardRect[2] &&
                itemCenterY <= artboardRect[1] && itemCenterY >= artboardRect[3]) {
                if (!artboardGroups[j]) {
                    artboardGroups[j] = [];
                    artboardOrder.push(j);
                }
                artboardGroups[j].push(selection[i]);
                assigned = true;
                break;
            }
        }

        if (!assigned) {
            noArtboardItems.push(selection[i]);
        }
    }

    for (var k = 0; k < artboardOrder.length; k++) {
        processBelowArtGroup(artboardGroups[artboardOrder[k]], settings, dimLayer);
    }

    if (noArtboardItems.length > 0) {
        processBelowArtGroup(noArtboardItems, settings, dimLayer);
    }
}

// For one group of items (a single artboard's worth, or unassigned items):
//   1. Add a Qty label under EACH item individually (like dimensions do)
//   2. Add ONE Scale label centered under the whole group, positioned below
//      whatever the lowest Qty/dim/art point ended up being
function processBelowArtGroup(items, settings, dimLayer) {
    var doc = app.activeDocument;
    var QTY_FONT_SIZE = 10;

    // Track the group's horizontal extent and the lowest Y reached after
    // all per-sign content has been placed. The Scale label uses both.
    var groupMinX = Number.MAX_VALUE;
    var groupMaxX = -Number.MAX_VALUE;
    var groupLowestY = Number.MAX_VALUE;

    // Per-sign pass: Qty label below each item
    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        var b = item.geometricBounds;
        var itemMinX = b[0];
        var itemMaxX = b[2];
        var itemMinY = b[3];
        var itemCenterX = (itemMinX + itemMaxX) / 2;

        if (itemMinX < groupMinX) groupMinX = itemMinX;
        if (itemMaxX > groupMaxX) groupMaxX = itemMaxX;

        // Where does the "stuff below this sign" start? If a bottom dim is
        // being added, start below that dim's text. Otherwise, start at the
        // art's bottom edge.
        var itemBottomY = itemMinY;
        if (settings.addWidth && settings.widthPos === "Below") {
            // Bottom dim text bottom = minY - offset - textOffset - fontSize (approx)
            itemBottomY = itemMinY - settings.offset - settings.textOffset - settings.fontSize;
        }

        if (settings.addQty) {
            // Same offset the dim line uses to clear the art -- keeps spacing
            // visually consistent with the dimension markers.
            var qtyTop = itemBottomY - settings.offset;

            var textFrame = dimLayer.textFrames.add();
            textFrame.contents = "Qty: " + settings.qtyValue;

            var characterStyle = getCharacterStyle("DimStyle");
            if (characterStyle) {
                textFrame.textRange.characterAttributes.characterStyle = characterStyle;
            }

            try {
                textFrame.textRange.characterAttributes.textFont = app.textFonts.getByName(settings.fontName);
            } catch (e) {}

            textFrame.textRange.characterAttributes.size = QTY_FONT_SIZE;
            textFrame.textRange.paragraphAttributes.justification = Justification.CENTER;

            textFrame.top = qtyTop;
            textFrame.left = itemCenterX - (textFrame.width / 2);

            var qtyBottom = qtyTop - textFrame.height;
            if (qtyBottom < groupLowestY) groupLowestY = qtyBottom;
        } else {
            // No Qty -- track the bottom of the art (or dim) for the Scale
            if (itemBottomY < groupLowestY) groupLowestY = itemBottomY;
        }
    }

    // One Scale label centered under the whole group, if requested
    if (settings.scaleTextPosition === "Below Centered") {
        var scaleCenterX = (groupMinX + groupMaxX) / 2;
        var scaleTop = groupLowestY - settings.offset;

        // Fall back to active layer if dim layer is locked/hidden
        var scaleLayer = dimLayer;
        try {
            if (scaleLayer.locked || !scaleLayer.visible) {
                scaleLayer = doc.activeLayer;
            }
        } catch (e) {}

        var scaleFrame = scaleLayer.textFrames.add();
        scaleFrame.contents = "Scale " + settings.scale;

        var characterStyle = getCharacterStyle("DimStyle");
        if (characterStyle) {
            scaleFrame.textRange.characterAttributes.characterStyle = characterStyle;
        }

        scaleFrame.textRange.characterAttributes.size = 7;
        scaleFrame.textRange.paragraphAttributes.justification = Justification.CENTER;

        scaleFrame.top = scaleTop;
        scaleFrame.left = scaleCenterX - (scaleFrame.width / 2);
    }
}

// Add scale text to all artboards that contain selected objects
// NOTE: "Below Centered" is now handled by addBelowArtTexts() so it can stack
// with the Qty label and reserve space for bottom dimensions. This function
// only handles "On Proof" and "Outside Proof".
function addScaleTextToArtboards(selection, settings, layer) {
    if (settings.scaleTextPosition === "None") return;
    if (settings.scaleTextPosition === "Below Centered") return;

    var doc = app.activeDocument;
    var artboardIndices = [];

    // For "Below Centered" option, check if art is on artboards
    if (settings.scaleTextPosition === "Below Centered") {
        var hasArtOnArtboard = false;
        
        // Find which artboards contain the selected objects
        for (var i = 0; i < selection.length; i++) {
            var itemBounds = selection[i].geometricBounds;
            var itemCenterX = (itemBounds[0] + itemBounds[2]) / 2;
            var itemCenterY = (itemBounds[1] + itemBounds[3]) / 2;
            
            // Check which artboard contains this object's center point
            for (var j = 0; j < doc.artboards.length; j++) {
                var artboardRect = doc.artboards[j].artboardRect;
                if (itemCenterX >= artboardRect[0] && itemCenterX <= artboardRect[2] &&
                    itemCenterY <= artboardRect[1] && itemCenterY >= artboardRect[3]) {
                    hasArtOnArtboard = true;
                    // Check if we haven't already added this artboard
                    var alreadyAdded = false;
                    for (var k = 0; k < artboardIndices.length; k++) {
                        if (artboardIndices[k] === j) {
                            alreadyAdded = true;
                            break;
                        }
                    }
                    if (!alreadyAdded) {
                        artboardIndices.push(j);
                    }
                    break;
                }
            }
        }
        
        // If no art is on artboards, add scale text without artboard reference
        if (!hasArtOnArtboard) {
            addScaleTextToArtboard(-1, settings, layer, selection);
            return;
        }
    } else {
        // For "On Proof" and "Outside Proof", find artboards normally
        for (var i = 0; i < selection.length; i++) {
            var itemBounds = selection[i].geometricBounds;
            var itemCenterX = (itemBounds[0] + itemBounds[2]) / 2;
            var itemCenterY = (itemBounds[1] + itemBounds[3]) / 2;
            
            // Check which artboard contains this object's center point
            for (var j = 0; j < doc.artboards.length; j++) {
                var artboardRect = doc.artboards[j].artboardRect;
                if (itemCenterX >= artboardRect[0] && itemCenterX <= artboardRect[2] &&
                    itemCenterY <= artboardRect[1] && itemCenterY >= artboardRect[3]) {
                    // Check if we haven't already added this artboard
                    var alreadyAdded = false;
                    for (var k = 0; k < artboardIndices.length; k++) {
                        if (artboardIndices[k] === j) {
                            alreadyAdded = true;
                            break;
                        }
                    }
                    if (!alreadyAdded) {
                        artboardIndices.push(j);
                    }
                    break;
                }
            }
        }
    }
    
    // Add scale text to each unique artboard
    for (var i = 0; i < artboardIndices.length; i++) {
        addScaleTextToArtboard(artboardIndices[i], settings, layer, selection);
    }
}

// Add scale text to a specific artboard
function addScaleTextToArtboard(artboardIndex, settings, layer, selection) {
    var doc = app.activeDocument;
    var x, y;
    
    if (settings.scaleTextPosition === "Below Centered") {
        // Calculate bounds for all selected objects
        var minX = Number.MAX_VALUE;
        var maxX = -Number.MAX_VALUE;
        var minY = Number.MAX_VALUE;
        
        if (artboardIndex === -1) {
            // No artboard - use all selected objects
            for (var i = 0; i < selection.length; i++) {
                var itemBounds = selection[i].geometricBounds;
                if (itemBounds[0] < minX) minX = itemBounds[0];
                if (itemBounds[2] > maxX) maxX = itemBounds[2];
                if (itemBounds[3] < minY) minY = itemBounds[3];
            }
        } else {
            // On artboard - use only objects on this artboard
            var artboard = doc.artboards[artboardIndex];
            var artboardRect = artboard.artboardRect;
            var artboardLeft = artboardRect[0];
            var artboardTop = artboardRect[1];
            var artboardRight = artboardRect[2];
            var artboardBottom = artboardRect[3];
            
            for (var i = 0; i < selection.length; i++) {
                var itemBounds = selection[i].geometricBounds;
                var itemCenterX = (itemBounds[0] + itemBounds[2]) / 2;
                var itemCenterY = (itemBounds[1] + itemBounds[3]) / 2;
                
                // Check if this item is on this artboard
                if (itemCenterX >= artboardLeft && itemCenterX <= artboardRight &&
                    itemCenterY <= artboardTop && itemCenterY >= artboardBottom) {
                    if (itemBounds[0] < minX) minX = itemBounds[0];
                    if (itemBounds[2] > maxX) maxX = itemBounds[2];
                    if (itemBounds[3] < minY) minY = itemBounds[3];
                }
            }
        }
        
        // Position 1 inch below the lowest point, centered horizontally
        x = (minX + maxX) / 2;
        y = minY - (1 * 72);
    } else {
        // "On Proof" and "Outside Proof" require an artboard
        var artboard = doc.artboards[artboardIndex];
        var artboardRect = artboard.artboardRect;
        var artboardLeft = artboardRect[0];
        var artboardTop = artboardRect[1];
        
        // Calculate position relative to artboard for proof positions
        x = artboardLeft + (4 * 72); // 4 inches from left edge
        
        if (settings.scaleTextPosition === "On Proof") {
            y = artboardTop - (6.9816 * 72);
        } else { // "Outside Proof"
            y = artboardTop - (8.625 * 72);
        }
    }
    
    // Create the scale text
    var scaleText = layer.textFrames.add();
    scaleText.contents = "Scale " + settings.scale;
    
    // Apply DimStyle character style
    var characterStyle = getCharacterStyle("DimStyle");
    if (characterStyle) {
        scaleText.textRange.characterAttributes.characterStyle = characterStyle;
    }
    
    // Set font size to 7pt
    scaleText.textRange.characterAttributes.size = 7;
    
    // Set center justification
    scaleText.textRange.paragraphAttributes.justification = Justification.CENTER;
    
    // Position the text (center it at the specified coordinates)
    scaleText.top = y;
    scaleText.left = x - (scaleText.width / 2);
}

// Show font customization dialog
function showFontCustomizationDialog(currentFontName, currentFontSize) {
    var fontDialog = new Window("dialog", "Font Settings");
    fontDialog.alignChildren = "fill";
    
    var fontPanel = fontDialog.add("panel", undefined, "Font Customization");
    fontPanel.alignChildren = "left";
    
    var fontRow = fontPanel.add("group");
    var fontLabel = fontRow.add("statictext", undefined, "Font:");
    fontLabel.preferredSize.width = 35;
    var fontDropdown = fontRow.add("dropdownlist", undefined, []);
    fontDropdown.preferredSize.width = 250;
    
    var sizeRow = fontPanel.add("group");
    var sizeLabel = sizeRow.add("statictext", undefined, "Size:");
    sizeLabel.preferredSize.width = 35;
    var fontSizeInput = sizeRow.add("edittext", undefined, currentFontSize.toString());
    fontSizeInput.characters = 4;
    
    // Populate font list
    var fonts = app.textFonts;
    var fontNames = [];
    for (var i = 0; i < fonts.length; i++) {
        fontNames.push(fonts[i].name);
    }
    fontNames.sort();
    
    for (var i = 0; i < fontNames.length; i++) {
        fontDropdown.add("item", fontNames[i]);
    }
    
    // Select the current font
    for (var i = 0; i < fontDropdown.items.length; i++) {
        if (fontDropdown.items[i].text === currentFontName) {
            fontDropdown.selection = i;
            break;
        }
    }
    if (!fontDropdown.selection && fontDropdown.items.length > 0) {
        fontDropdown.selection = 0;
    }
    
    // Buttons
    var buttonGroup = fontDialog.add("group");
    buttonGroup.add("button", undefined, "OK", {name: "ok"});
    buttonGroup.add("button", undefined, "Cancel", {name: "cancel"});
    
    if (fontDialog.show() == 1) {
        return {
            fontName: fontDropdown.selection.text,
            fontSize: parseFloat(fontSizeInput.text) || 12
        };
    }
    
    return null;
}

function main() {
    if (app.documents.length === 0) {
        alert("Please open a document first.");
        return;
    }
    
    var doc = app.activeDocument;
    var sel = doc.selection;
    
    if (sel.length === 0) {
        // Nothing selected: scan the Dimensions layer for measurement text
        // and offer to switch between decimal and fraction formats.
        convertDimensionTexts();
        return;
    }

    // Remember the originally selected art so it can be reselected at the end,
    // letting other scripts keep working with the same objects.
    var originalSelection = [];
    for (var si = 0; si < sel.length; si++) {
        originalSelection.push(sel[si]);
    }

    // Default font settings
    var currentFontName = "MyriadPro-Semibold";
    var currentFontSize = 12;
    
    // Try to find default font in available fonts
    var fonts = app.textFonts;
    var defaultFound = false;
    for (var i = 0; i < fonts.length; i++) {
        if (fonts[i].name === currentFontName) {
            defaultFound = true;
            break;
        }
    }
    
    // If default not found, try alternate naming
    if (!defaultFound) {
        for (var i = 0; i < fonts.length; i++) {
            if (fonts[i].name.indexOf("Myriad") > -1 && 
                fonts[i].name.indexOf("Semibold") > -1) {
                currentFontName = fonts[i].name;
                defaultFound = true;
                break;
            }
        }
    }
    
    // If still not found, use first available font
    if (!defaultFound && fonts.length > 0) {
        currentFontName = fonts[0].name;
    }
    
    // Detect existing scale on artboards with selected objects
    var existingScale = detectExistingScale(sel);
    
    // Create configuration dialog
    var dialog = new Window("dialog", "Auto-Dimension Tool");
    dialog.alignChildren = "fill";
    
    // Scale and Precision settings
    var scalePanel = dialog.add("panel", undefined, "Drawing Settings");
    scalePanel.alignChildren = "left";
    
    var scaleRow = scalePanel.add("group");
    var precLabel = scaleRow.add("statictext", undefined, "Precision:");
    precLabel.preferredSize.width = 60;
    var precisionDropdown = scaleRow.add("dropdownlist", undefined, ["0", "0.0", "0.00", "0.000"]);
    precisionDropdown.selection = 3; // Default to 0.000
    
    var scaleLabel = scaleRow.add("statictext", undefined, "Scale:");
    scaleLabel.preferredSize.width = 48;
    var scaleDropdown = scaleRow.add("dropdownlist", undefined, ["1:1", "1:2", "1:4", "1:5", "1:8", "1:10", "1:16", "1:20", "1:40", "1:50", "1:100", "2:1", "4:1", "8:1", "10:1", "100:2", "Custom"]);
    
    // If existing scale detected, try to select it; otherwise default to 1:10
    if (existingScale) {
        var scaleFound = false;
        for (var i = 0; i < scaleDropdown.items.length; i++) {
            if (scaleDropdown.items[i].text === existingScale) {
                scaleDropdown.selection = i;
                scaleFound = true;
                break;
            }
        }
        if (!scaleFound) {
            scaleDropdown.selection = 5; // Default to 1:10
        }
    } else {
        scaleDropdown.selection = 5; // Default to 1:10
    }
    
    scaleDropdown.preferredSize.width = 80;
    
    // Parse existing scale for custom values
    var initialNumerator = "1";
    var initialDenominator = "1";
    if (existingScale && existingScale.indexOf(":") > -1) {
        var scaleParts = existingScale.split(":");
        if (scaleParts.length === 2) {
            initialNumerator = scaleParts[0];
            initialDenominator = scaleParts[1];
            
            // If it's a custom scale (not in dropdown), select Custom and populate fields
            var scaleFound = false;
            for (var i = 0; i < scaleDropdown.items.length - 1; i++) { // -1 to exclude "Custom"
                if (scaleDropdown.items[i].text === existingScale) {
                    scaleFound = true;
                    break;
                }
            }
            if (!scaleFound) {
                scaleDropdown.selection = scaleDropdown.items.length - 1; // Select "Custom"
            }
        }
    }
    
    var customScaleNumerator = scaleRow.add("edittext", undefined, initialNumerator);
    customScaleNumerator.characters = 3;
    customScaleNumerator.enabled = true; // Always enabled - typing auto-switches to Custom

    scaleRow.add("statictext", undefined, ":");

    var customScaleDenominator = scaleRow.add("edittext", undefined, initialDenominator);
    customScaleDenominator.characters = 3;
    customScaleDenominator.enabled = true; // Always enabled - typing auto-switches to Custom

    // Flag to suppress field-change -> Custom switching when fields are updated programmatically
    var suppressCustomSwitch = false;

    // If the initial dropdown selection is a preset, sync the fields to it so they
    // always reflect the current scale (gives users a starting point if they want to edit)
    (function syncFieldsToDropdown() {
        if (!scaleDropdown.selection) return;
        var selText = scaleDropdown.selection.text;
        if (selText === "Custom") return;
        var parts = selText.split(":");
        if (parts.length !== 2) return;
        suppressCustomSwitch = true;
        customScaleNumerator.text = parts[0];
        customScaleDenominator.text = parts[1];
        suppressCustomSwitch = false;
    })();

    // When the dropdown changes to a preset, update the custom fields to match
    scaleDropdown.onChange = function() {
        if (!scaleDropdown.selection) return;
        var selText = scaleDropdown.selection.text;
        if (selText === "Custom") return;
        var parts = selText.split(":");
        if (parts.length !== 2) return;
        suppressCustomSwitch = true;
        customScaleNumerator.text = parts[0];
        customScaleDenominator.text = parts[1];
        suppressCustomSwitch = false;
    };

    // When user types in either custom field, auto-switch the dropdown to "Custom"
    function switchToCustomScale() {
        if (suppressCustomSwitch) return;
        if (scaleDropdown.selection && scaleDropdown.selection.text === "Custom") return;
        for (var i = 0; i < scaleDropdown.items.length; i++) {
            if (scaleDropdown.items[i].text === "Custom") {
                scaleDropdown.selection = i;
                break;
            }
        }
    }
    customScaleNumerator.onChanging = switchToCustomScale;
    customScaleDenominator.onChanging = switchToCustomScale;
    
    // Dimensions to add with position checkboxes
    var dimPanel = dialog.add("panel", undefined, "Dimension Placement");
    dimPanel.alignChildren = "center";
    
    // Visual diagram with checkboxes
    var topRow = dimPanel.add("group");
    topRow.alignment = "center";
    var checkTop = topRow.add("checkbox", undefined, "Top");
    checkTop.value = true; // Default on
    
    var middleRow = dimPanel.add("group");
    var checkLeft = middleRow.add("checkbox", undefined, "Left");
    checkLeft.value = true; // Default on
    var spacer = middleRow.add("statictext", undefined, "          ");
    spacer.preferredSize.width = 80;
    var checkRight = middleRow.add("checkbox", undefined, "Right");
    checkRight.value = false; // Default off
    
    var bottomRow = dimPanel.add("group");
    bottomRow.alignment = "center";
    var checkBottom = bottomRow.add("checkbox", undefined, "Bottom");
    checkBottom.value = false; // Default off
    
    // Add mutual exclusivity for left/right
    checkLeft.onClick = function() {
        if (checkLeft.value) {
            checkRight.value = false;
        }
    };
    
    checkRight.onClick = function() {
        if (checkRight.value) {
            checkLeft.value = false;
        }
    };
    
    // Add mutual exclusivity for top/bottom
    checkTop.onClick = function() {
        if (checkTop.value) {
            checkBottom.value = false;
        }
    };
    
    checkBottom.onClick = function() {
        if (checkBottom.value) {
            checkTop.value = false;
        }
    };
    
    // Rounding - row of selectable buttons (radio buttons) for one-click selection
    var roundingGroup = dimPanel.add("group");
    roundingGroup.add("statictext", undefined, "Rounding:");
    var roundingValues = ["None", "1/16\" (0.0625)", "1/8\" (0.125)", "1/4\" (0.25)", "1/2\" (0.5)", "Integer"];
    var roundingLabels = ["None", "1/16\"", "1/8\"", "1/4\"", "1/2\"", "Integer"];
    var roundingRadios = [];
    for (var ri = 0; ri < roundingLabels.length; ri++) {
        var rRb = roundingGroup.add("radiobutton", undefined, roundingLabels[ri]);
        roundingRadios.push(rRb);
    }
    roundingRadios[1].value = true; // Default to 1/16"

    // Format - decimal vs fraction display for dimension text.
    // Fractions need a rounding increment, so "None" rounding forces Decimal.
    var formatGroup = dimPanel.add("group");
    formatGroup.add("statictext", undefined, "Format:");
    var formatDecimalRadio = formatGroup.add("radiobutton", undefined, "Decimal");
    var formatFractionRadio = formatGroup.add("radiobutton", undefined, "Fractions");
    formatFractionRadio.value = true; // Default to fractions

    // Remember the user's format choice while rounding is "None" so it can be
    // restored if they pick a rounding increment again
    var formatBeforeNone = "Fractions";
    function syncFormatToRounding() {
        if (getRoundingValue() === "None") {
            if (formatFractionRadio.enabled) {
                formatBeforeNone = formatFractionRadio.value ? "Fractions" : "Decimal";
            }
            formatFractionRadio.value = false;
            formatDecimalRadio.value = true;
            formatFractionRadio.enabled = false;
        } else if (!formatFractionRadio.enabled) {
            formatFractionRadio.enabled = true;
            if (formatBeforeNone === "Fractions") {
                formatFractionRadio.value = true;
                formatDecimalRadio.value = false;
            }
        }
    }
    for (var fri = 0; fri < roundingRadios.length; fri++) {
        roundingRadios[fri].onClick = syncFormatToRounding;
    }
    syncFormatToRounding();

    // Include Scale - row of selectable buttons (radio buttons) for one-click selection
    var scaleTextGroup = dimPanel.add("group");
    scaleTextGroup.add("statictext", undefined, "Include Scale:");
    var scaleTextValues = ["None", "On Proof", "Outside Proof", "Below Centered"];
    var scaleTextRadios = [];
    for (var si = 0; si < scaleTextValues.length; si++) {
        var sRb = scaleTextGroup.add("radiobutton", undefined, scaleTextValues[si]);
        scaleTextRadios.push(sRb);
    }

    // If existing scale detected, default to "None"; otherwise default to "On Proof"
    if (existingScale) {
        scaleTextRadios[0].value = true; // "None"
    } else {
        scaleTextRadios[1].value = true; // "On Proof"
    }

    // Helpers to read which button is selected
    function getRoundingValue() {
        for (var i = 0; i < roundingRadios.length; i++) {
            if (roundingRadios[i].value) return roundingValues[i];
        }
        return roundingValues[1];
    }
    function getScaleTextValue() {
        for (var i = 0; i < scaleTextRadios.length; i++) {
            if (scaleTextRadios[i].value) return scaleTextValues[i];
        }
        return scaleTextValues[1];
    }
    
    // Quantity Label (goes on the Dimensions layer, with the dimensions)
    // Default value "XX" acts as a placeholder the user can double-click in
    // Illustrator to replace per-sign. If quantities are all the same, the
    // user can pre-fill the real number here and skip that step entirely.
    var qtyPanel = dialog.add("panel", undefined, "Quantity Label");
    qtyPanel.alignChildren = "left";
    var qtyRow = qtyPanel.add("group");
    var checkQty = qtyRow.add("checkbox", undefined, "Add \"Qty:\" label below each sign");
    checkQty.value = true;
    var qtyValueLabel = qtyRow.add("statictext", undefined, "   Qty:");
    qtyValueLabel.preferredSize.width = 40;
    var qtyValueInput = qtyRow.add("edittext", undefined, "XX");
    qtyValueInput.characters = 5;

    // Style settings (combined graphic style and text)
    var stylePanel = dialog.add("panel", undefined, "Style");
    stylePanel.alignChildren = "left";
    
    var colorGroup = stylePanel.add("group");
    colorGroup.add("statictext", undefined, "Color:");
    var colorDropdown = colorGroup.add("dropdownlist", undefined, ["Black", "White", "Red"]);
    colorDropdown.selection = 0;
    
    // Font customization button
    var fontButtonGroup = stylePanel.add("group");
    var fontInfoText = fontButtonGroup.add("statictext", undefined, "Font: " + currentFontName + " (" + currentFontSize + "pt)");
    fontInfoText.preferredSize.width = 250;
    var fontCustomizeButton = fontButtonGroup.add("button", undefined, "Customize...");
    
    fontCustomizeButton.onClick = function() {
        var result = showFontCustomizationDialog(currentFontName, currentFontSize);
        if (result) {
            currentFontName = result.fontName;
            currentFontSize = result.fontSize;
            fontInfoText.text = "Font: " + currentFontName + " (" + currentFontSize + "pt)";
        }
    };
    
    // Collision detection
    var collisionPanel = dialog.add("panel", undefined, "Smart Placement");
    collisionPanel.alignChildren = "left";
    var checkCollision = collisionPanel.add("checkbox", undefined, "Skip dimensions that would overlap other objects");
    checkCollision.value = true;
    
    // Layer settings
    var layerPanel = dialog.add("panel", undefined, "Layer Options");
    layerPanel.alignChildren = "left";
    var checkNewLayer = layerPanel.add("checkbox", undefined, "Create dimensions on new layer");
    checkNewLayer.value = true;
    var layerNameGroup = layerPanel.add("group");
    layerNameGroup.add("statictext", undefined, "Layer name:");
    var layerNameInput = layerNameGroup.add("edittext", undefined, "Dimensions");
    layerNameInput.characters = 15;
    
    // Buttons
    var buttonGroup = dialog.add("group");
    buttonGroup.add("button", undefined, "OK", {name: "ok"});
    buttonGroup.add("button", undefined, "Cancel", {name: "cancel"});
    
    // Show dialog
    if (dialog.show() == 2) return;
    
    // Get settings
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
    
    // Get rounding increment
    var roundingIncrement;
    var roundingText = getRoundingValue();
    if (roundingText === "None") {
        roundingIncrement = 0;
    } else if (roundingText.indexOf("0.0625") > -1) {
        roundingIncrement = 0.0625;
    } else if (roundingText.indexOf("0.125") > -1) {
        roundingIncrement = 0.125;
    } else if (roundingText.indexOf("0.25") > -1) {
        roundingIncrement = 0.25;
    } else if (roundingText.indexOf("0.5") > -1) {
        roundingIncrement = 0.5;
    } else { // Integer
        roundingIncrement = 1.0;
    }
    
    var settings = {
        addWidth: checkTop.value || checkBottom.value,
        addHeight: checkLeft.value || checkRight.value,
        widthPos: checkTop.value ? "Above" : "Below",
        heightPos: checkLeft.value ? "Left" : "Right",
        offset: 0.1 * 72,
        scale: scaleText,
        scaleRatio: scaleRatio,
        precision: precisionDropdown.selection.index,
        roundingIncrement: roundingIncrement,
        useFractions: formatFractionRadio.value,
        scaleTextPosition: getScaleTextValue(),
        lineColor: colorDropdown.selection.text,
        lineWeight: 1,
        extensionLength: 0.083 * 72,
        extensionOffset: 0.046 * 72,
        fontName: currentFontName,
        fontSize: currentFontSize,
        textOffset: 0.05 * 72,
        checkCollision: checkCollision.value,
        useNewLayer: checkNewLayer.value,
        layerName: layerNameInput.text || "Dimensions",
        useArrowheads: true, // Default to true, will be set to false if styles missing
        addQty: checkQty.value,
        qtyValue: (qtyValueInput.text && qtyValueInput.text.replace(/^\s+|\s+$/g, "").length > 0) ? qtyValueInput.text.replace(/^\s+|\s+$/g, "") : "XX"
    };
    
    if (!checkTop.value && !checkBottom.value && !checkLeft.value && !checkRight.value) {
        alert("Please select at least one dimension location.");
        return;
    }

    // Map user-friendly color names to graphic style names
    var colorToStyleMap = {
        "Black": "Dim-Blk",
        "White": "Dim-Wht",
        "Red": "Dim-Red"
    };
    settings.graphicStyleName = colorToStyleMap[settings.lineColor];

    // Check if the required graphic style exists
    var styleExists = getGraphicStyle(settings.graphicStyleName) !== null;

    if (!styleExists) {
        var confirmDialog = confirm(
            "The graphic style '" + settings.graphicStyleName + "' was not found in this document.\n\n" +
            "Dimensions will be created WITHOUT arrowheads.\n\n" +
            "Click OK to continue without arrowheads, or Cancel to stop.\n\n" +
            "(To use arrowheads, you need to have the dimension graphic styles in your document.)"
        );

        if (!confirmDialog) {
            return; // User cancelled
        }

        settings.useArrowheads = false;
    }
    
    // Create or get dimension layer
    var dimLayer;
    if (settings.useNewLayer) {
        dimLayer = getOrCreateLayer(settings.layerName);
    } else {
        dimLayer = doc.activeLayer;
    }
    
    // Get all object bounds for collision detection using the new method
    var allBounds = [];
    for (var i = 0; i < sel.length; i++) {
        var simplifiedBounds = getSimplifiedBounds(sel[i], null);
        allBounds.push(simplifiedBounds);
    }
    
    // Process each selected object
    var processedCount = 0;
    var skippedCount = 0;
    
    for (var i = 0; i < sel.length; i++) {
        try {
            var result = addDimensionsToObject(sel[i], settings, dimLayer, allBounds, i);
            if (result.added) processedCount++;
            if (result.skipped) skippedCount += result.skipped;
        } catch (e) {
            // Skip objects that can't be dimensioned
        }
    }
    
    var message = "Dimensions added to " + processedCount + " object(s).";
    if (skippedCount > 0) {
        message += "\n" + skippedCount + " dimension(s) skipped due to overlaps.";
    }
    
    // Add below-art texts (Qty label and/or "Below Centered" Scale).
    // These get stacked together so they don't overlap each other or the
    // bottom dimension.
    if (settings.addQty || settings.scaleTextPosition === "Below Centered") {
        addBelowArtTexts(sel, settings, dimLayer);
    }

    // Add scale text for the artboard-anchored positions ("On Proof" /
    // "Outside Proof"). "Below Centered" is handled above.
    if (settings.scaleTextPosition === "On Proof" || settings.scaleTextPosition === "Outside Proof") {
        addScaleTextToArtboards(sel, settings, dimLayer);
    }

    // Reselect the original art so downstream scripts can continue with it.
    try {
        doc.selection = null;
        for (var rs = 0; rs < originalSelection.length; rs++) {
            try { originalSelection[rs].selected = true; } catch (eOne) {}
        }
        app.redraw();
    } catch (eRestore) {}

    // (Completion confirmation removed -- runs silently.)
}

// NEW METHOD: Get bounds using temporary artboard and Fit to Selected Art
function getSimplifiedBounds(obj, tempLayer) {
    var doc = app.activeDocument;
    var originalArtboardIndex = -1;
    var artboardCountBefore = doc.artboards.length;
    
    try {
        // Store the current artboard index
        originalArtboardIndex = doc.artboards.getActiveArtboardIndex();
        
        // Create a temporary artboard (just needs to exist, size doesn't matter initially)
        var tempArtboard = doc.artboards.add([0, 0, 100, -100]);
        var tempArtboardIndex = doc.artboards.length - 1;
        
        // Set the temp artboard as active
        doc.artboards.setActiveArtboardIndex(tempArtboardIndex);
        
        // Select only the current object
        doc.selection = null;
        obj.selected = true;
        
        // Use Fit Artboard to Selected Art menu command
        // This will resize the artboard to match the selected object's bounds
        app.executeMenuCommand("Fit Artboard to selected Art");
        
        // Refresh to ensure artboard has updated
        app.redraw();
        
        // Get the artboard rect AFTER the fit command (this is now the bounds of our object)
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
            // Try to remove temp artboard if it exists
            if (doc.artboards.length > artboardCountBefore) {
                doc.artboards.remove(doc.artboards.length - 1);
            }
        } catch (cleanupError) {}
        
        try {
            // Restore original artboard if needed
            if (originalArtboardIndex >= 0 && originalArtboardIndex < doc.artboards.length) {
                doc.artboards.setActiveArtboardIndex(originalArtboardIndex);
            }
        } catch (cleanupError) {}
        
        return obj.geometricBounds;
    }
}



// Get or create layer
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

// Check if two rectangles overlap
function boundsOverlap(bounds1, bounds2, buffer) {
    buffer = buffer || 0;
    return !(bounds1[2] + buffer < bounds2[0] ||
             bounds1[0] - buffer > bounds2[2] ||
             bounds1[3] + buffer > bounds2[1] ||
             bounds1[1] - buffer < bounds2[3]);
}

// Check if dimension would collide with any objects
function wouldCollide(dimBounds, allBounds, currentIndex, settings) {
    if (!settings.checkCollision) return false;
    
    var buffer = 5;
    
    for (var i = 0; i < allBounds.length; i++) {
        if (i === currentIndex) continue;
        if (boundsOverlap(dimBounds, allBounds[i], buffer)) {
            return true;
        }
    }
    return false;
}

// Add dimensions to a single object
function addDimensionsToObject(obj, settings, layer, allBounds, currentIndex) {
    var bounds = allBounds[currentIndex]; // Use the pre-calculated simplified bounds
    
    var width = bounds[2] - bounds[0];
    var height = bounds[1] - bounds[3];
    
    // Convert to inches and apply scale
    var widthInches = (width / 72) * settings.scaleRatio;
    var heightInches = (height / 72) * settings.scaleRatio;
    
    var added = false;
    var skipped = 0;
    
    // Add width dimension
    if (settings.addWidth) {
        var widthY;
        var widthDimBounds;
        
        if (settings.widthPos === "Above") {
            widthY = bounds[1] + settings.offset;
            widthDimBounds = [
                bounds[0] - settings.offset, 
                widthY + settings.extensionLength + settings.textOffset + settings.fontSize,
                bounds[2] + settings.offset,
                widthY - settings.extensionOffset
            ];
        } else {
            widthY = bounds[3] - settings.offset;
            widthDimBounds = [
                bounds[0] - settings.offset,
                widthY + settings.extensionOffset,
                bounds[2] + settings.offset,
                widthY - settings.extensionLength - settings.textOffset - settings.fontSize
            ];
        }
        
        if (!wouldCollide(widthDimBounds, allBounds, currentIndex, settings)) {
            createDimension(
                bounds[0], widthY,
                bounds[2], widthY,
                widthInches,
                true,
                settings.widthPos,
                settings,
                layer
            );
            added = true;
        } else {
            skipped++;
        }
    }
    
    // Add height dimension
    if (settings.addHeight) {
        var heightX;
        var heightDimBounds;
        
        if (settings.heightPos === "Right") {
            heightX = bounds[2] + settings.offset;
            heightDimBounds = [
                heightX - settings.extensionOffset,
                bounds[1] + settings.offset,
                heightX + settings.extensionLength + settings.textOffset + settings.fontSize,
                bounds[3] - settings.offset
            ];
        } else {
            heightX = bounds[0] - settings.offset;
            heightDimBounds = [
                heightX - settings.extensionLength - settings.textOffset - settings.fontSize,
                bounds[1] + settings.offset,
                heightX + settings.extensionOffset,
                bounds[3] - settings.offset
            ];
        }
        
        if (!wouldCollide(heightDimBounds, allBounds, currentIndex, settings)) {
            createDimension(
                heightX, bounds[1],
                heightX, bounds[3],
                heightInches,
                false,
                settings.heightPos,
                settings,
                layer
            );
            added = true;
        } else {
            skipped++;
        }
    }
    
    return { added: added, skipped: skipped };
}

// Apply dimension line styling directly (with arrowheads)
function applyDimensionStyling(line, styleName) {
    // Determine color based on style name
    var strokeColor;
    if (styleName === "Dim-Wht") {
        strokeColor = new CMYKColor();
        strokeColor.cyan = 0;
        strokeColor.magenta = 0;
        strokeColor.yellow = 0;
        strokeColor.black = 0;
    } else if (styleName === "Dim-Red") {
        strokeColor = new CMYKColor();
        strokeColor.cyan = 0;
        strokeColor.magenta = 100;
        strokeColor.yellow = 100;
        strokeColor.black = 0;
    } else { // Dim-Blk
        strokeColor = new CMYKColor();
        strokeColor.cyan = 0;
        strokeColor.magenta = 0;
        strokeColor.yellow = 0;
        strokeColor.black = 100;
    }
    
    // Apply basic stroke properties
    line.stroked = true;
    line.filled = false;
    line.strokeWidth = 1;
    line.strokeColor = strokeColor;
    
    // Try to apply arrowheads - wrap in try/catch to not break if it fails
    try {
        // Check if arrowhead properties exist
        if (typeof line.strokeArrowStart !== 'undefined') {
            var arrowStart = app.arrowheadsStartList;
            var arrowEnd = app.arrowheadsEndList;
            
            // Arrow 7 is at index 6
            if (arrowStart && arrowStart.length > 6) {
                line.strokeArrowStart = arrowStart[6];
                line.strokeArrowStartScale = 70;
            }
            if (arrowEnd && arrowEnd.length > 6) {
                line.strokeArrowEnd = arrowEnd[6];
                line.strokeArrowEndScale = 70;
            }
        }
    } catch (e) {
        // Silently continue if arrowheads fail - line will still be created
    }
    
    return line;
}

// Get graphic style by name, or return null if not found
function getGraphicStyle(styleName) {
    var doc = app.activeDocument;
    try {
        for (var i = 0; i < doc.graphicStyles.length; i++) {
            if (doc.graphicStyles[i].name === styleName) {
                return doc.graphicStyles[i];
            }
        }
    } catch (e) {}
    return null;
}

// Get text color based on graphic style name
function getTextColor(styleName) {
    var color;
    if (styleName === "Dim-Wht") {
        color = new CMYKColor();
        color.cyan = 0;
        color.magenta = 0;
        color.yellow = 0;
        color.black = 0;
    } else if (styleName === "Dim-Red") {
        color = new CMYKColor();
        color.cyan = 0;
        color.magenta = 100;
        color.yellow = 100;
        color.black = 0;
    } else { // Dim-Blk
        color = new CMYKColor();
        color.cyan = 0;
        color.magenta = 0;
        color.yellow = 0;
        color.black = 100;
    }
    return color;
}

// Create a dimension line with text
function createDimension(x1, y1, x2, y2, value, isHorizontal, position, settings, layer) {
    var doc = app.activeDocument;
    
    try {
        var dimGroup = layer.groupItems.add();
        dimGroup.name = "Dimension";
        
        // Get the graphic style if it exists
        var graphicStyle = getGraphicStyle(settings.graphicStyleName);
        
        // Get color based on graphic style for both text and extension lines
        var color = getTextColor(settings.graphicStyleName);
        
        // Create extension lines (plain lines without arrowheads)
        var extLine1, extLine2;
        if (isHorizontal) {
            if (position === "Above") {
                extLine1 = createPlainLine(x1, y1 - settings.extensionOffset, x1, y1 + settings.extensionLength, color, dimGroup);
                extLine2 = createPlainLine(x2, y1 - settings.extensionOffset, x2, y1 + settings.extensionLength, color, dimGroup);
            } else {
                extLine1 = createPlainLine(x1, y1 + settings.extensionOffset, x1, y1 - settings.extensionLength, color, dimGroup);
                extLine2 = createPlainLine(x2, y1 + settings.extensionOffset, x2, y1 - settings.extensionLength, color, dimGroup);
            }
        } else {
            if (position === "Right") {
                extLine1 = createPlainLine(x1 - settings.extensionOffset, y1, x1 + settings.extensionLength, y1, color, dimGroup);
                extLine2 = createPlainLine(x1 - settings.extensionOffset, y2, x1 + settings.extensionLength, y2, color, dimGroup);
            } else {
                extLine1 = createPlainLine(x1 + settings.extensionOffset, y1, x1 - settings.extensionLength, y1, color, dimGroup);
                extLine2 = createPlainLine(x1 + settings.extensionOffset, y2, x1 - settings.extensionLength, y2, color, dimGroup);
            }
        }
        
        // Create dimension line (this will get arrowheads)
        var dimLine = createSimpleLine(x1, y1, x2, y2, dimGroup);
        
        // Apply graphic style if it exists and arrowheads are enabled, otherwise apply styling directly
        if (settings.useArrowheads && graphicStyle) {
            try {
                graphicStyle.applyTo(dimLine);
            } catch (e) {
                // If style application fails, apply basic styling without arrowheads
                dimLine.stroked = true;
                dimLine.filled = false;
                dimLine.strokeWidth = 1;
                dimLine.strokeColor = color;
            }
        } else {
            // No arrowheads - just apply basic line styling
            dimLine.stroked = true;
            dimLine.filled = false;
            dimLine.strokeWidth = 1;
            dimLine.strokeColor = color;
        }
        
        // Create text with matching color
        var text = dimGroup.textFrames.add();
        
        // Round and format per settings (fraction or decimal)
        var formattedValue = formatDimensionValue(value, settings);
        text.contents = formattedValue + '"';
        
        try {
            text.textRange.characterAttributes.textFont = app.textFonts.getByName(settings.fontName);
        } catch (e) {
            // Use default font if specified font not found
        }
        
        text.textRange.characterAttributes.size = settings.fontSize;
        text.textRange.characterAttributes.fillColor = color;

        // Set text to center alignment for better editability
        text.textRange.paragraphAttributes.justification = Justification.CENTER;

        // Style the fraction part typographically (raised numerator, fraction
        // slash, small denominator) so values like 11 11/16" read cleanly
        if (settings.useFractions && formattedValue.indexOf("/") > -1) {
            styleFractionText(text, settings.fontSize);
        }
        
        // Position text based on whether it's horizontal or vertical
        var centerX = (x1 + x2) / 2;
        var centerY = (y1 + y2) / 2;
        
        if (isHorizontal) {
            // Horizontal dimension - center horizontally over the line
            text.left = centerX - (text.width / 2);
            
            // Position above or below the dimension line with consistent offset measurement
            if (position === "Below") {
                // Below: text top edge is offset distance below the line
                text.top = y1 - settings.textOffset;
            } else { // "Above"
                // Above: text bottom edge is offset distance above the line
                text.top = y1 + settings.textOffset + text.height;
            }
        } else {
            // Vertical dimension - rotate first, then position
            // Rotate the text 90 degrees
            text.rotate(90);
            
            // After rotation, width and height are swapped
            // Center along the dimension line
            text.top = centerY + (text.height / 2);
            
            // Position left or right with consistent offset measurement
            if (position === "Left") {
                // Left: text right edge is offset distance left of the line
                text.left = x1 - settings.textOffset - text.width;
            } else { // "Right"
                // Right: text left edge is offset distance right of the line
                text.left = x1 + settings.textOffset;
            }
        }
    } catch (e) {
        // If anything fails in dimension creation, log it but don't stop the script
        // alert("Error creating dimension: " + e.toString());
    }
}

// Create a simple line without styling (for dimension line that will get graphic style)
function createSimpleLine(x1, y1, x2, y2, parent) {
    var line = parent.pathItems.add();
    line.stroked = true;
    line.filled = false;
    line.setEntirePath([[x1, y1], [x2, y2]]);
    return line;
}

// Create a plain line with color but no arrowheads (for extension lines)
function createPlainLine(x1, y1, x2, y2, color, parent) {
    var line = parent.pathItems.add();
    line.stroked = true;
    line.strokeColor = color;
    line.strokeWidth = 1;
    line.filled = false;
    line.setEntirePath([[x1, y1], [x2, y2]]);
    return line;
}

// Run the script
main();
