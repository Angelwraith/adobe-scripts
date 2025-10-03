/*@METADATA{
  "name": "Image Quality Analysis",
  "description": "Analyze selected images and label them with resolution info",
  "version": "1.0",
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
    
    // Show scale selection dialog
    var scaleDialog = new Window("dialog", "Image Quality Analysis - Select Scale");
    scaleDialog.alignChildren = "fill";
    scaleDialog.spacing = 15;
    scaleDialog.margins = 20;
    
    var titleText = scaleDialog.add("statictext", undefined, "Select the drawing scale:");
    titleText.graphics.font = ScriptUI.newFont("dialog", "Bold", 12);
    
    var scaleGroup = scaleDialog.add("group");
    scaleGroup.add("statictext", undefined, "Scale:");
    var scaleDropdown = scaleGroup.add("dropdownlist", undefined, ["1:1", "1:2", "1:4", "1:5", "1:8", "1:10", "1:16", "1:20", "1:40", "1:50", "1:100", "2:1", "4:1", "8:1", "10:1", "100:2", "Custom"]);
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
    
    // Process each clipping mask
    var labelsCreated = 0;
    for (var i = 0; i < clippingMasks.length; i++) {
        var clippedGroup = clippingMasks[i];
        
        // Get the bounds of the clipping path (first item) BEFORE processing
        var clippingPath = clippedGroup.pageItems[0];
        var maskBounds = clippingPath.geometricBounds;
        
        var result = processClippingMask(clippedGroup, maskBounds, analysisLayer, scaleRatio, scaleFactor);
        if (result) labelsCreated++;
    }
    
    if (labelsCreated === 0) {
        alert("No clipping masks with images found in selection.");
    } else {
        alert("Analysis complete!\n" + labelsCreated + " label(s) created.");
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
        
        if (item.typename === "GroupItem" && item.clipped) {
            // This is a clipping mask - add it
            results.push(item);
        }
        
        // Also check inside groups for nested clipping masks
        if (item.typename === "GroupItem" && item.pageItems && item.pageItems.length > 0) {
            collectClippingMasks(item.pageItems, results);
        }
    }
}

function processClippingMask(clippedGroup, maskBounds, layer, scaleRatio, scaleFactor) {
    try {
        // maskBounds is already passed in - it's the clipping path bounds
        
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
            // Use the maskBounds that was passed in
            createImageLabel(maskBounds, estimatedPPI, isLowRes, layer);
            return true;
        }
        
    } catch (e) {
        alert("Error in processClippingMask: " + e.toString());
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
    
    // Set font
    try {
        label.textRange.characterAttributes.textFont = app.textFonts.getByName("MyriadPro-Bold");
    } catch (e) {
        try {
            label.textRange.characterAttributes.textFont = app.textFonts.getByName("Myriad-Bold");
        } catch (e2) {}
    }
    
    label.textRange.characterAttributes.size = 18;
    
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

// Run the script
main();
