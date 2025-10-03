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

// Add scale text to all artboards that contain selected objects
function addScaleTextToArtboards(selection, settings, layer) {
    if (settings.scaleTextPosition === "None") return;
    
    var doc = app.activeDocument;
    var artboardIndices = [];
    
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
    
    // Add scale text to each unique artboard
    for (var i = 0; i < artboardIndices.length; i++) {
        addScaleTextToArtboard(artboardIndices[i], settings, layer);
    }
}

// Add scale text to a specific artboard
function addScaleTextToArtboard(artboardIndex, settings, layer) {
    var doc = app.activeDocument;
    var artboard = doc.artboards[artboardIndex];
    var artboardRect = artboard.artboardRect;
    var artboardLeft = artboardRect[0];
    var artboardTop = artboardRect[1];
    
    // Calculate position relative to artboard
    var x = artboardLeft + (15 * 72); // 15 inches from left edge
    var y;
    
    if (settings.scaleTextPosition === "On Proof") {
        y = artboardTop - (9.75 * 72);
    } else { // "Outside Proof"
        y = artboardTop - (11 * 72);
    }
    
    // Create the scale text
    var scaleText = layer.textFrames.add();
    scaleText.contents = "Scale " + settings.scale;
    
    // Apply DimStyle character style
    var characterStyle = getCharacterStyle("DimStyle");
    if (characterStyle) {
        scaleText.textRange.characterAttributes.characterStyle = characterStyle;
    }
    
    // Set font size to 14pt
    scaleText.textRange.characterAttributes.size = 14;
    
    // Set center justification
    scaleText.textRange.paragraphAttributes.justification = Justification.CENTER;
    
    // Position the text (center it at the specified coordinates)
    scaleText.top = y;
    scaleText.left = x - (scaleText.width / 2);
}// Illustrator Script: Auto-Dimension Tool v2
// Automatically add dimension annotations with smart collision detection

/*@METADATA{
  "name": "Smart Dimension Tool",
  "description": "Add dimensions to an array of signs without the fuss",
  "version": "1.0",
  "target": "illustrator",
  "tags": ["Measure", "Smart", "Utility"]
}@END_METADATA*/

function main() {
    if (app.documents.length === 0) {
        alert("Please open a document first.");
        return;
    }
    
    var doc = app.activeDocument;
    var sel = doc.selection;
    
    if (sel.length === 0) {
        alert("Please select one or more objects to dimension.");
        return;
    }
    
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
    scaleDropdown.selection = 5; // Default to 1:10
    scaleDropdown.preferredSize.width = 80;
    
    var customScaleNumerator = scaleRow.add("edittext", undefined, "1");
    customScaleNumerator.characters = 3;
    customScaleNumerator.enabled = false;
    
    scaleRow.add("statictext", undefined, ":");
    
    var customScaleDenominator = scaleRow.add("edittext", undefined, "1");
    customScaleDenominator.characters = 3;
    customScaleDenominator.enabled = false;
    
    // Enable/disable custom scale inputs based on dropdown selection
    scaleDropdown.onChange = function() {
        var isCustom = scaleDropdown.selection.text === "Custom";
        customScaleNumerator.enabled = isCustom;
        customScaleDenominator.enabled = isCustom;
    };
    
    // Dimensions to add with position dropdowns inline
    var dimPanel = dialog.add("panel", undefined, "Dimensions to Add");
    dimPanel.alignChildren = "left";
    
    var dimRow = dimPanel.add("group");
    var checkWidth = dimRow.add("checkbox", undefined, "Add Width");
    checkWidth.value = true;
    var widthPosDropdown = dimRow.add("dropdownlist", undefined, ["Above", "Below"]);
    widthPosDropdown.selection = 0;
    
    dimRow.add("statictext", undefined, "     ");
    
    var checkHeight = dimRow.add("checkbox", undefined, "Add Height");
    checkHeight.value = true;
    var heightPosDropdown = dimRow.add("dropdownlist", undefined, ["Left", "Right"]);
    heightPosDropdown.selection = 1;
    
    var offsetGroup = dimPanel.add("group");
    offsetGroup.add("statictext", undefined, "Offset from object (inches):");
    var offsetInput = offsetGroup.add("edittext", undefined, "0.125");
    offsetInput.characters = 6;
    
    var roundingGroup = dimPanel.add("group");
    roundingGroup.add("statictext", undefined, "Rounding:");
    var roundingDropdown = roundingGroup.add("dropdownlist", undefined, ["None", "1/16\" (0.0625)", "1/8\" (0.125)", "1/4\" (0.25)", "1/2\" (0.5)", "Integer"]);
    roundingDropdown.selection = 1; // Default to 1/16"
    
    var scaleTextGroup = dimPanel.add("group");
    scaleTextGroup.add("statictext", undefined, "Include Scale:");
    var scaleTextDropdown = scaleTextGroup.add("dropdownlist", undefined, ["None", "On Proof", "Outside Proof"]);
    scaleTextDropdown.selection = 1; // Default to "On Proof"
    
    // Style settings (combined graphic style and text)
    var stylePanel = dialog.add("panel", undefined, "Style");
    stylePanel.alignChildren = "left";
    
    var colorGroup = stylePanel.add("group");
    colorGroup.add("statictext", undefined, "Color:");
    var colorDropdown = colorGroup.add("dropdownlist", undefined, ["Black", "White", "Red"]);
    colorDropdown.selection = 0;
    
    var fontRow = stylePanel.add("group");
    var fontLabel = fontRow.add("statictext", undefined, "Font:");
    fontLabel.preferredSize.width = 35;
    var fontDropdown = fontRow.add("dropdownlist", undefined, []);
    fontDropdown.preferredSize.width = 180;
    
    var sizeLabel = fontRow.add("statictext", undefined, "Size:");
    sizeLabel.preferredSize.width = 32;
    var fontSizeInput = fontRow.add("edittext", undefined, "18");
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
    
    // Try to select Myriad Pro Semibold as default
    var defaultFont = "MyriadPro-Semibold";
    for (var i = 0; i < fontDropdown.items.length; i++) {
        if (fontDropdown.items[i].text === defaultFont) {
            fontDropdown.selection = i;
            break;
        }
    }
    // If not found, try alternate naming
    if (!fontDropdown.selection) {
        for (var i = 0; i < fontDropdown.items.length; i++) {
            if (fontDropdown.items[i].text.indexOf("Myriad") > -1 && 
                fontDropdown.items[i].text.indexOf("Semibold") > -1) {
                fontDropdown.selection = i;
                break;
            }
        }
    }
    if (!fontDropdown.selection) {
        fontDropdown.selection = 0;
    }
    
    var textOffsetGroup = stylePanel.add("group");
    textOffsetGroup.add("statictext", undefined, "Text Offset from Line (inches):");
    var textOffsetInput = textOffsetGroup.add("edittext", undefined, "0.125");
    textOffsetInput.characters = 6;
    
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
        scaleRatio = denom / num; // For 1:10, this gives 10/1 = 10
    } else {
        scaleText = scaleDropdown.selection.text;
        var parts = scaleText.split(":");
        scaleRatio = parseFloat(parts[1]) / parseFloat(parts[0]); // For 1:10, this gives 10/1 = 10
    }
    
    // Get rounding increment
    var roundingIncrement;
    var roundingText = roundingDropdown.selection.text;
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
        addWidth: checkWidth.value,
        addHeight: checkHeight.value,
        widthPos: widthPosDropdown.selection.text,
        heightPos: heightPosDropdown.selection.text,
        offset: parseFloat(offsetInput.text) * 72 || 9, // Convert inches to points
        scale: scaleText,
        scaleRatio: scaleRatio, // The multiplier to apply to dimensions
        precision: precisionDropdown.selection.index, // 0, 1, 2, or 3 decimal places
        roundingIncrement: roundingIncrement, // Rounding increment in inches
        scaleTextPosition: scaleTextDropdown.selection.text, // None, On Proof, or Outside Proof
        lineColor: colorDropdown.selection.text, // Black, White, or Red (user-friendly names)
        lineWeight: 1,
        extensionLength: 0.125 * 72, // Fixed at 0.125 inches
        extensionOffset: 0.069 * 72, // Fixed at 0.069 inches
        fontName: fontDropdown.selection.text,
        fontSize: parseFloat(fontSizeInput.text) || 18,
        textOffset: parseFloat(textOffsetInput.text) * 72 || 9, // Convert inches to points
        checkCollision: checkCollision.value,
        useNewLayer: checkNewLayer.value,
        layerName: layerNameInput.text || "Dimensions"
    };
    
    // Map user-friendly color names to graphic style names
    var colorToStyleMap = {
        "Black": "Dim-Blk",
        "White": "Dim-Wht",
        "Red": "Dim-Red"
    };
    settings.graphicStyleName = colorToStyleMap[settings.lineColor];
    
    if (!settings.addWidth && !settings.addHeight) {
        alert("Please select at least one dimension type to add.");
        return;
    }
    
    // Create or get dimension layer
    var dimLayer;
    if (settings.useNewLayer) {
        dimLayer = getOrCreateLayer(settings.layerName);
    } else {
        dimLayer = doc.activeLayer;
    }
    
    // Get all object bounds for collision detection
    var allBounds = [];
    for (var i = 0; i < sel.length; i++) {
        allBounds.push(sel[i].geometricBounds);
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
    
    // Add scale text if requested
    if (settings.scaleTextPosition !== "None") {
        addScaleTextToArtboards(sel, settings, dimLayer);
    }
    
    alert(message);
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
    return !(bounds1[2] + buffer < bounds2[0] || // right < left
             bounds1[0] - buffer > bounds2[2] || // left > right
             bounds1[3] + buffer > bounds2[1] || // bottom > top
             bounds1[1] - buffer < bounds2[3]);  // top < bottom
}

// Check if dimension would collide with any objects
function wouldCollide(dimBounds, allBounds, currentIndex, settings) {
    if (!settings.checkCollision) return false;
    
    // Add a larger buffer for better collision detection
    var buffer = 5;
    
    for (var i = 0; i < allBounds.length; i++) {
        if (i === currentIndex) continue;
        if (boundsOverlap(dimBounds, allBounds[i], buffer)) {
            return true;
        }
    }
    return false;
}

// Find the actual visible bounds of an object, accounting for nested clipping masks
// Based on getVisibleBounds by Josh Duncan (joshbduncan.com)
function getVisibleBounds(object) {
    var bounds, clippedItem;
    if (object.typename == "GroupItem") {
        // if the object is clipped
        if (object.clipped) {
            // check all sub objects to find the clipping path
            for (var i = 0; i < object.pageItems.length; i++) {
                if (object.pageItems[i].clipping) {
                    clippedItem = object.pageItems[i];
                    break;
                } else if (object.pageItems[i].typename == "CompoundPathItem") {
                    if (object.pageItems[i].pathItems[0].clipping) {
                        clippedItem = object.pageItems[i];
                        break;
                    }
                } else {
                    clippedItem = object.pageItems[i];
                    break;
                }
            }
            bounds = clippedItem.geometricBounds;
        } else {
            // if the object is not clipped, check for clipped children first
            var hasClippedChildren = false;
            for (var i = 0; i < object.pageItems.length; i++) {
                if (object.pageItems[i].typename == "GroupItem" && object.pageItems[i].clipped) {
                    // Found a clipped child - use its bounds instead of calculating from all children
                    bounds = getVisibleBounds(object.pageItems[i]);
                    hasClippedChildren = true;
                    break;
                }
            }
            
            if (!hasClippedChildren) {
                // if the object is not clipped and has no clipped children
                var subObjectBounds;
                var allBoundPoints = [[], [], [], []];
                // get the bounds of every object in the group
                for (var i = 0; i < object.pageItems.length; i++) {
                    subObjectBounds = getVisibleBounds(object.pageItems[i]);
                    allBoundPoints[0].push(subObjectBounds[0]);
                    allBoundPoints[1].push(subObjectBounds[1]);
                    allBoundPoints[2].push(subObjectBounds[2]);
                    allBoundPoints[3].push(subObjectBounds[3]);
                }
                // determine the groups bounds from it sub object bound points
                bounds = [
                    Math.min.apply(Math, allBoundPoints[0]),
                    Math.max.apply(Math, allBoundPoints[1]),
                    Math.max.apply(Math, allBoundPoints[2]),
                    Math.min.apply(Math, allBoundPoints[3]),
                ];
            }
        }
    } else {
        bounds = object.geometricBounds;
    }
    return bounds;
}

// Add dimensions to a single object
function addDimensionsToObject(obj, settings, layer, allBounds, currentIndex) {
    var bounds = getVisibleBounds(obj);
    
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
                true, // horizontal
                settings.widthPos, // pass position setting
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
                false, // vertical
                settings.heightPos, // pass position setting
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

// Get graphic style by name
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
        color = new GrayColor();
        color.gray = 0; // White text = 0 gray value
    } else if (styleName === "Dim-Red") {
        color = new CMYKColor();
        color.cyan = 0;
        color.magenta = 100;
        color.yellow = 100;
        color.black = 0;
    } else { // Dim-Blk
        color = new GrayColor();
        color.gray = 100; // Black text = 100 gray value
    }
    return color;
}

// Create a dimension line with text
function createDimension(x1, y1, x2, y2, value, isHorizontal, position, settings, layer) {
    var doc = app.activeDocument;
    var dimGroup = layer.groupItems.add();
    dimGroup.name = "Dimension";
    
    // Get the graphic style
    var graphicStyle = getGraphicStyle(settings.graphicStyleName);
    if (!graphicStyle) {
        alert("Graphic style '" + settings.graphicStyleName + "' not found. Please create it first.");
        return;
    }
    
    // Get color based on graphic style for both text and extension lines
    var color = getTextColor(settings.graphicStyleName);
    
    // Create extension lines (plain lines without arrowheads)
    var extLine1, extLine2;
    if (isHorizontal) {
        if (position === "Above") {
            // Above: extension lines go up from the object
            extLine1 = createPlainLine(x1, y1 - settings.extensionOffset, x1, y1 + settings.extensionLength, color, dimGroup);
            extLine2 = createPlainLine(x2, y1 - settings.extensionOffset, x2, y1 + settings.extensionLength, color, dimGroup);
        } else {
            // Below: extension lines go down from the object
            extLine1 = createPlainLine(x1, y1 + settings.extensionOffset, x1, y1 - settings.extensionLength, color, dimGroup);
            extLine2 = createPlainLine(x2, y1 + settings.extensionOffset, x2, y1 - settings.extensionLength, color, dimGroup);
        }
    } else {
        if (position === "Right") {
            // Right: extension lines go right from the object
            extLine1 = createPlainLine(x1 - settings.extensionOffset, y1, x1 + settings.extensionLength, y1, color, dimGroup);
            extLine2 = createPlainLine(x1 - settings.extensionOffset, y2, x1 + settings.extensionLength, y2, color, dimGroup);
        } else {
            // Left: extension lines go left from the object
            extLine1 = createPlainLine(x1 + settings.extensionOffset, y1, x1 - settings.extensionLength, y1, color, dimGroup);
            extLine2 = createPlainLine(x1 + settings.extensionOffset, y2, x1 - settings.extensionLength, y2, color, dimGroup);
        }
    }
    
    // Create dimension line (this will get arrowheads from graphic style)
    var dimLine = createSimpleLine(x1, y1, x2, y2, dimGroup);
    
    // Apply graphic style to dimension line only (this has the arrowheads)
    graphicStyle.applyTo(dimLine);
    
    // Create text with matching color
    var text = dimGroup.textFrames.add();
    
    // Round to the specified increment if not "None"
    var roundedValue = value;
    if (settings.roundingIncrement > 0) {
        roundedValue = Math.round(value / settings.roundingIncrement) * settings.roundingIncrement;
    }
    
    // Format the value with precision, then remove trailing zeros
    var formattedValue = roundedValue.toFixed(settings.precision);
    // Remove trailing zeros after decimal point
    if (formattedValue.indexOf('.') > -1) {
        formattedValue = formattedValue.replace(/\.?0+$/, '');
    }
    
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