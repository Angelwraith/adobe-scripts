/*@METADATA{
  "name": "Smart Dimension Tool",
  "description": "Add dimensions to an array of signs without the fuss",
  "version": "3.3",
  "target": "illustrator",
  "tags": ["Measure", "Smart", "Utility"]
}@END_METADATA*/

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
            fontSize: parseFloat(fontSizeInput.text) || 18
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
        alert("Please select one or more objects to dimension.");
        return;
    }
    
    // Default font settings
    var currentFontName = "MyriadPro-Semibold";
    var currentFontSize = 18;
    
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
    heightPosDropdown.selection = 0; // Default to Left
    
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
        scaleRatio = denom / num;
    } else {
        scaleText = scaleDropdown.selection.text;
        var parts = scaleText.split(":");
        scaleRatio = parseFloat(parts[1]) / parseFloat(parts[0]);
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
        offset: parseFloat(offsetInput.text) * 72 || 9,
        scale: scaleText,
        scaleRatio: scaleRatio,
        precision: precisionDropdown.selection.index,
        roundingIncrement: roundingIncrement,
        scaleTextPosition: scaleTextDropdown.selection.text,
        lineColor: colorDropdown.selection.text,
        lineWeight: 1,
        extensionLength: 0.125 * 72,
        extensionOffset: 0.069 * 72,
        fontName: currentFontName,
        fontSize: currentFontSize,
        textOffset: parseFloat(textOffsetInput.text) * 72 || 9,
        checkCollision: checkCollision.value,
        useNewLayer: checkNewLayer.value,
        layerName: layerNameInput.text || "Dimensions",
        useArrowheads: true // Default to true, will be set to false if styles missing
    };
    
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
    
    // Add scale text if requested
    if (settings.scaleTextPosition !== "None") {
        addScaleTextToArtboards(sel, settings, dimLayer);
    }
    
    alert(message);
}

// NEW METHOD: Get bounds using temporary artboard and Fit to Selected Art
function getSimplifiedBounds(obj, tempLayer) {
    var doc = app.activeDocument;
    
    try {
        // Store the current artboard index
        var originalArtboardIndex = doc.artboards.getActiveArtboardIndex();
        
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
            if (doc.artboards.length > 0) {
                var lastIndex = doc.artboards.length - 1;
                doc.artboards.remove(lastIndex);
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
        
        // Round to the specified increment if not "None"
        var roundedValue = value;
        if (settings.roundingIncrement > 0) {
            roundedValue = Math.round(value / settings.roundingIncrement) * settings.roundingIncrement;
        }
        
        // Format the value with precision, then remove trailing zeros
        var formattedValue = roundedValue.toFixed(settings.precision);
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