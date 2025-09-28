/*
@METADATA
{
  "name": "Cut Path Separator",
  "description": "Organize Cut/Print Data Into Respective Layers",
  "version": "1.0",
  "target": "illustrator",
  "tags": ["Cut", "Path", "Separator", "processors"]
}
@END_METADATA
*/


(function() {
    // Check if there's an active document
    if (app.documents.length === 0) {
        alert("Please open a document first.");
        return;
    }
    
    var doc = app.activeDocument;
    var spotColorNames = ["CutThrough2-Outside", "CutThrough1-Inside", "CutThrough-Knifecut", "CutContour", "Spot1"];
    var totalMovedPaths = 0;
    var targetStrokeWidth = 5 / doc.scaleFactor; // Adjust for Large Canvas documents
    
    // Function to check if two paths are duplicates (same path, same location)
    function pathsAreDuplicates(path1, path2) {
        try {
            // Check if paths have the same number of path points
            if (path1.pathPoints.length !== path2.pathPoints.length) {
                return false;
            }
            
            // Check if geometric bounds are the same (within tolerance)
            var bounds1 = path1.geometricBounds;
            var bounds2 = path2.geometricBounds;
            var tolerance = 0.1; // Small tolerance for floating point comparison
            
            for (var i = 0; i < 4; i++) {
                if (Math.abs(bounds1[i] - bounds2[i]) > tolerance) {
                    return false;
                }
            }
            
            // Check if path points are in the same positions
            for (var i = 0; i < path1.pathPoints.length; i++) {
                var point1 = path1.pathPoints[i];
                var point2 = path2.pathPoints[i];
                
                // Compare anchor points
                if (Math.abs(point1.anchor[0] - point2.anchor[0]) > tolerance ||
                    Math.abs(point1.anchor[1] - point2.anchor[1]) > tolerance) {
                    return false;
                }
                
                // Compare left direction points
                if (Math.abs(point1.leftDirection[0] - point2.leftDirection[0]) > tolerance ||
                    Math.abs(point1.leftDirection[1] - point2.leftDirection[1]) > tolerance) {
                    return false;
                }
                
                // Compare right direction points
                if (Math.abs(point1.rightDirection[0] - point2.rightDirection[0]) > tolerance ||
                    Math.abs(point1.rightDirection[1] - point2.rightDirection[1]) > tolerance) {
                    return false;
                }
            }
            
            return true; // Paths are duplicates (regardless of color)
            
        } catch (e) {
            // If there's an error comparing paths, assume they're not duplicates
            return false;
        }
    }
    
    // Function to check for duplicate paths in a layer
    function checkForDuplicatePaths(layer) {
        var allPaths = [];
        
        // Collect all path items in the layer (including in groups)
        function collectPaths(container) {
            for (var i = 0; i < container.pageItems.length; i++) {
                var item = container.pageItems[i];
                if (item.typename === "PathItem") {
                    allPaths.push(item);
                } else if (item.typename === "GroupItem") {
                    collectPaths(item);
                } else if (item.typename === "CompoundPathItem") {
                    for (var j = 0; j < item.pathItems.length; j++) {
                        allPaths.push(item.pathItems[j]);
                    }
                }
            }
        }
        
        collectPaths(layer);
        
        // Compare paths for duplicates (colors don't need to match)
        for (var i = 0; i < allPaths.length - 1; i++) {
            for (var j = i + 1; j < allPaths.length; j++) {
                if (pathsAreDuplicates(allPaths[i], allPaths[j])) {
                    return true; // Found duplicates
                }
            }
        }
        
        return false; // No duplicates found
    }
    
    // Function to enable overprint fill for all items in Spot1 layer
    function enableOverprintFillSpot1(layer) {
        var allItems = [];
        
        // Collect all path items and text frames in the layer (including in groups)
        function collectItems(container) {
            for (var i = 0; i < container.pageItems.length; i++) {
                var item = container.pageItems[i];
                if (item.typename === "PathItem" || item.typename === "TextFrame") {
                    allItems.push(item);
                } else if (item.typename === "GroupItem") {
                    collectItems(item);
                } else if (item.typename === "CompoundPathItem") {
                    for (var j = 0; j < item.pathItems.length; j++) {
                        allItems.push(item.pathItems[j]);
                    }
                }
            }
        }
        
        collectItems(layer);
        
        // Enable overprint fill for all items that don't have it enabled
        for (var i = 0; i < allItems.length; i++) {
            var item = allItems[i];
            try {
                if (item.typename === "PathItem") {
                    // For path items, enable fillOverprint if filled and not already enabled
                    if (item.filled && item.fillOverprint === false) {
                        item.fillOverprint = true;
                    }
                } else if (item.typename === "TextFrame") {
                    // For text frames, try different overprint approaches
                    try {
                        // Select the text frame first
                        app.activeDocument.selection = [item];
                        
                        // Try accessing through textRange
                        var textRange = item.textRange;
                        
                        // Method 1: Set overprint on the entire text range
                        try {
                            textRange.characterAttributes.overprintFill = true;
                        } catch (e1) {
                            // Method 2: Try fillOverprint instead
                            try {
                                textRange.characterAttributes.fillOverprint = true;
                            } catch (e2) {
                                // Method 3: Try setting on individual characters
                                try {
                                    for (var k = 0; k < textRange.characters.length; k++) {
                                        var character = textRange.characters[k];
                                        try {
                                            character.characterAttributes.overprintFill = true;
                                        } catch (charError1) {
                                            try {
                                                character.characterAttributes.fillOverprint = true;
                                            } catch (charError2) {
                                                // Continue with next character
                                                continue;
                                            }
                                        }
                                    }
                                } catch (e3) {
                                    // Method 4: Try direct on text frame
                                    try {
                                        item.overprintFill = true;
                                    } catch (e4) {
                                        // All methods failed
                                        continue;
                                    }
                                }
                            }
                        }
                        
                    } catch (textError) {
                        // Continue with next item if text processing fails
                        continue;
                    }
                }
            } catch (e) {
                // Continue with other items if one fails
                continue;
            }
        }
        
        // Clear selection when done
        try {
            app.activeDocument.selection = [];
        } catch (clearError) {
            // Ignore if clearing selection fails
        }
    }
    
    // Function to find the spot color in the document
    function findSpotColor(colorName) {
        for (var i = 0; i < doc.spots.length; i++) {
            if (doc.spots[i].name === colorName) {
                return doc.spots[i];
            }
        }
        return null;
    }
    
    // Function to check if a path item uses the target spot color
    function usesSpotColor(pathItem, spotColor) {
        if (!pathItem || !spotColor) return false;
        
        // Check fill color
        if (pathItem.filled && pathItem.fillColor.typename === "SpotColor") {
            if (pathItem.fillColor.spot.name === spotColor.name) {
                return true;
            }
        }
        
        // Check stroke color
        if (pathItem.stroked && pathItem.strokeColor.typename === "SpotColor") {
            if (pathItem.strokeColor.spot.name === spotColor.name) {
                return true;
            }
        }
        
        return false;
    }
    
    // Function to check if a text frame uses the target spot color
    function usesSpotColorText(textFrame, spotColor) {
        if (!textFrame || !spotColor) return false;
        
        // Check text fill color
        try {
            var textRange = textFrame.textRange;
            var fillColor = textRange.fillColor;
            
            if (fillColor && fillColor.typename === "SpotColor") {
                if (fillColor.spot.name === spotColor.name) {
                    return true;
                }
            }
            
            // Check text stroke color
            var strokeColor = textRange.strokeColor;
            if (strokeColor && strokeColor.typename === "SpotColor") {
                if (strokeColor.spot.name === spotColor.name) {
                    return true;
                }
            }
        } catch (e) {
            // If there's an error accessing text properties, skip this text frame
            return false;
        }
        
        return false;
    }
    
    // Function to recursively search through all items in a container
    function searchForPaths(container, spotColor) {
        var foundPaths = [];
        
        for (var i = 0; i < container.pageItems.length; i++) {
            var item = container.pageItems[i];
            
            // If it's a path item, check if it uses our spot color
            if (item.typename === "PathItem" && usesSpotColor(item, spotColor)) {
                foundPaths.push(item);
            }
            // If it's a group or compound path, search recursively
            else if (item.typename === "GroupItem") {
                var groupPaths = searchForPaths(item, spotColor);
                foundPaths = foundPaths.concat(groupPaths);
            }
            else if (item.typename === "CompoundPathItem") {
                // Check each path in the compound path
                for (var j = 0; j < item.pathItems.length; j++) {
                    if (usesSpotColor(item.pathItems[j], spotColor)) {
                        foundPaths.push(item);
                        break; // Only add the compound path once
                    }
                }
            }
        }
        
        return foundPaths;
    }
    
    // Function to search for both paths and text items for Spot1 color only
    function searchForPathsAndText(container, spotColor, includeText) {
        var foundItems = [];
        
        for (var i = 0; i < container.pageItems.length; i++) {
            var item = container.pageItems[i];
            
            // If it's a path item, check if it uses our spot color
            if (item.typename === "PathItem" && usesSpotColor(item, spotColor)) {
                foundItems.push(item);
            }
            // If includeText is true and it's a text item, check if it uses our spot color
            else if (includeText && item.typename === "TextFrame" && usesSpotColorText(item, spotColor)) {
                foundItems.push(item);
            }
            // If it's a group or compound path, search recursively
            else if (item.typename === "GroupItem") {
                var groupItems = searchForPathsAndText(item, spotColor, includeText);
                foundItems = foundItems.concat(groupItems);
            }
            else if (item.typename === "CompoundPathItem") {
                // Check each path in the compound path
                for (var j = 0; j < item.pathItems.length; j++) {
                    if (usesSpotColor(item.pathItems[j], spotColor)) {
                        foundItems.push(item);
                        break; // Only add the compound path once
                    }
                }
            }
        }
        
        return foundItems;
    }
    
    // Function to create or find a layer with the given name
    function getOrCreateLayer(layerName) {
        // First, check if layer already exists
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === layerName) {
                return doc.layers[i];
            }
        }
        
        // Find the REG layer to position our new layer under it
        var regLayer = null;
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === "REG") {
                regLayer = doc.layers[i];
                break;
            }
        }
        
        // Create the new layer
        var newLayer = doc.layers.add();
        newLayer.name = layerName;
        
        // Move the new layer under the REG layer if found
        if (regLayer) {
            newLayer.move(regLayer, ElementPlacement.PLACEAFTER);
        }
        
        return newLayer;
    }
    
    // Function to standardize stroke weights in a layer
    function standardizeStrokeWeights(layer) {
        var allPaths = [];
        
        // Collect all path items in the layer (including in groups)
        function collectPaths(container) {
            for (var i = 0; i < container.pageItems.length; i++) {
                var item = container.pageItems[i];
                if (item.typename === "PathItem") {
                    allPaths.push(item);
                } else if (item.typename === "GroupItem") {
                    collectPaths(item);
                } else if (item.typename === "CompoundPathItem") {
                    for (var j = 0; j < item.pathItems.length; j++) {
                        allPaths.push(item.pathItems[j]);
                    }
                }
            }
        }
        
        collectPaths(layer);
        
        // Check and fix stroke weights for stroked paths only
        for (var i = 0; i < allPaths.length; i++) {
            var path = allPaths[i];
            if (path.stroked && Math.abs(path.strokeWidth - targetStrokeWidth) > 0.01) {
                path.strokeWidth = targetStrokeWidth;
            }
        }
    }
    
    // Function to swap fill and stroke of a path item
    function swapFillAndStroke(pathItem) {
        // Store current fill properties
        var tempFilled = pathItem.filled;
        var tempFillColor = null;
        if (pathItem.filled) {
            // Create a copy of the fill color
            if (pathItem.fillColor.typename === "SpotColor") {
                tempFillColor = new SpotColor();
                tempFillColor.spot = pathItem.fillColor.spot;
                tempFillColor.tint = pathItem.fillColor.tint;
            } else if (pathItem.fillColor.typename === "CMYKColor") {
                tempFillColor = new CMYKColor();
                tempFillColor.cyan = pathItem.fillColor.cyan;
                tempFillColor.magenta = pathItem.fillColor.magenta;
                tempFillColor.yellow = pathItem.fillColor.yellow;
                tempFillColor.black = pathItem.fillColor.black;
            } else if (pathItem.fillColor.typename === "RGBColor") {
                tempFillColor = new RGBColor();
                tempFillColor.red = pathItem.fillColor.red;
                tempFillColor.green = pathItem.fillColor.green;
                tempFillColor.blue = pathItem.fillColor.blue;
            }
        }
        
        // Store current stroke properties
        var tempStroked = pathItem.stroked;
        var tempStrokeColor = null;
        var tempStrokeWidth = pathItem.strokeWidth;
        if (pathItem.stroked) {
            // Create a copy of the stroke color
            if (pathItem.strokeColor.typename === "SpotColor") {
                tempStrokeColor = new SpotColor();
                tempStrokeColor.spot = pathItem.strokeColor.spot;
                tempStrokeColor.tint = pathItem.strokeColor.tint;
            } else if (pathItem.strokeColor.typename === "CMYKColor") {
                tempStrokeColor = new CMYKColor();
                tempStrokeColor.cyan = pathItem.strokeColor.cyan;
                tempStrokeColor.magenta = pathItem.strokeColor.magenta;
                tempStrokeColor.yellow = pathItem.strokeColor.yellow;
                tempStrokeColor.black = pathItem.strokeColor.black;
            } else if (pathItem.strokeColor.typename === "RGBColor") {
                tempStrokeColor = new RGBColor();
                tempStrokeColor.red = pathItem.strokeColor.red;
                tempStrokeColor.green = pathItem.strokeColor.green;
                tempStrokeColor.blue = pathItem.strokeColor.blue;
            }
        }
        
        // Apply stroke properties to fill
        pathItem.filled = tempStroked;
        if (tempStroked && tempStrokeColor) {
            pathItem.fillColor = tempStrokeColor;
        }
        
        // Apply fill properties to stroke
        pathItem.stroked = tempFilled;
        if (tempFilled && tempFillColor) {
            pathItem.strokeColor = tempFillColor;
        }
        
        // Set stroke width to target weight if now stroked
        if (pathItem.stroked) {
            pathItem.strokeWidth = targetStrokeWidth;
        }
    }
    
    // Function to process knife layer specifically
    function processKnifeLayer(layer, knifeSpotColor) {
        var allPaths = [];
        
        // Collect all path items in the layer (including in groups)
        function collectPaths(container) {
            for (var i = 0; i < container.pageItems.length; i++) {
                var item = container.pageItems[i];
                if (item.typename === "PathItem") {
                    allPaths.push(item);
                } else if (item.typename === "GroupItem") {
                    collectPaths(item);
                } else if (item.typename === "CompoundPathItem") {
                    for (var j = 0; j < item.pathItems.length; j++) {
                        allPaths.push(item.pathItems[j]);
                    }
                }
            }
        }
        
        collectPaths(layer);
        
        // Check each path for CutThrough-Knifecut fill color and swap if found
        for (var i = 0; i < allPaths.length; i++) {
            var path = allPaths[i];
            if (path.filled && 
                path.fillColor.typename === "SpotColor" && 
                path.fillColor.spot.name === knifeSpotColor.name) {
                swapFillAndStroke(path);
            }
        }
    }
    
    // Function to process CutThrough2-Outside layer specifically
    function processCutThrough2Layer(layer, cutThrough2SpotColor) {
        var allPaths = [];
        
        // Collect all path items in the layer (including in groups)
        function collectPaths(container) {
            for (var i = 0; i < container.pageItems.length; i++) {
                var item = container.pageItems[i];
                if (item.typename === "PathItem") {
                    allPaths.push(item);
                } else if (item.typename === "GroupItem") {
                    collectPaths(item);
                } else if (item.typename === "CompoundPathItem") {
                    for (var j = 0; j < item.pathItems.length; j++) {
                        allPaths.push(item.pathItems[j]);
                    }
                }
            }
        }
        
        collectPaths(layer);
        
        // Check each path for CutThrough2-Outside fill color and no stroke, then swap
        for (var i = 0; i < allPaths.length; i++) {
            var path = allPaths[i];
            if (path.filled && 
                !path.stroked &&
                path.fillColor.typename === "SpotColor" && 
                path.fillColor.spot.name === cutThrough2SpotColor.name) {
                swapFillAndStroke(path);
            }
        }
    }
    
    // Function to get appearance properties of a path item
    function getPathAppearance(pathItem) {
        var appearance = {
            fillType: pathItem.filled ? pathItem.fillColor.typename : "none",
            fillColor: "",
            strokeType: pathItem.stroked ? pathItem.strokeColor.typename : "none",
            strokeColor: "",
            strokeWidth: pathItem.stroked ? pathItem.strokeWidth : 0
        };
        
        // Get fill color details
        if (pathItem.filled) {
            if (pathItem.fillColor.typename === "SpotColor") {
                appearance.fillColor = pathItem.fillColor.spot.name;
            } else if (pathItem.fillColor.typename === "CMYKColor") {
                appearance.fillColor = "C:" + pathItem.fillColor.cyan + 
                                    " M:" + pathItem.fillColor.magenta + 
                                    " Y:" + pathItem.fillColor.yellow + 
                                    " K:" + pathItem.fillColor.black;
            } else if (pathItem.fillColor.typename === "RGBColor") {
                appearance.fillColor = "R:" + pathItem.fillColor.red + 
                                    " G:" + pathItem.fillColor.green + 
                                    " B:" + pathItem.fillColor.blue;
            }
        }
        
        // Get stroke color details
        if (pathItem.stroked) {
            if (pathItem.strokeColor.typename === "SpotColor") {
                appearance.strokeColor = pathItem.strokeColor.spot.name;
            } else if (pathItem.strokeColor.typename === "CMYKColor") {
                appearance.strokeColor = "C:" + pathItem.strokeColor.cyan + 
                                      " M:" + pathItem.strokeColor.magenta + 
                                      " Y:" + pathItem.strokeColor.yellow + 
                                      " K:" + pathItem.strokeColor.black;
            } else if (pathItem.strokeColor.typename === "RGBColor") {
                appearance.strokeColor = "R:" + pathItem.strokeColor.red + 
                                      " G:" + pathItem.strokeColor.green + 
                                      " B:" + pathItem.strokeColor.blue;
            }
        }
        
        return appearance;
    }
    
    // Function to compare two appearance objects
    function appearancesMatch(app1, app2) {
        return (app1.fillType === app2.fillType &&
                app1.fillColor === app2.fillColor &&
                app1.strokeType === app2.strokeType &&
                app1.strokeColor === app2.strokeColor &&
                Math.abs(app1.strokeWidth - app2.strokeWidth) < 0.01 * doc.scaleFactor); // Allow small floating point differences, adjusted for scale
    }
    
    // Function to check appearance consistency in a layer
    function checkLayerAppearanceConsistency(layer) {
        var allPaths = [];
        
        // Collect all path items in the layer (including in groups)
        function collectPaths(container) {
            for (var i = 0; i < container.pageItems.length; i++) {
                var item = container.pageItems[i];
                if (item.typename === "PathItem") {
                    allPaths.push(item);
                } else if (item.typename === "GroupItem") {
                    collectPaths(item);
                } else if (item.typename === "CompoundPathItem") {
                    for (var j = 0; j < item.pathItems.length; j++) {
                        allPaths.push(item.pathItems[j]);
                    }
                }
            }
        }
        
        collectPaths(layer);
        
        if (allPaths.length <= 1) {
            return true; // Consistent if 0 or 1 paths
        }
        
        // Compare all paths against the first path's appearance
        var referenceAppearance = getPathAppearance(allPaths[0]);
        
        for (var i = 1; i < allPaths.length; i++) {
            var currentAppearance = getPathAppearance(allPaths[i]);
            if (!appearancesMatch(referenceAppearance, currentAppearance)) {
                return false;
            }
        }
        
        return true;
    }
    
    try {
        var processedLayers = []; // Keep track of layers we created/used
        
        // Process each spot color in order
        for (var colorIndex = 0; colorIndex < spotColorNames.length; colorIndex++) {
            var currentSpotColorName = spotColorNames[colorIndex];
            var targetSpotColor = findSpotColor(currentSpotColorName);
            var pathsToMove = [];
            
            // Skip this color if it doesn't exist
            if (!targetSpotColor) {
                continue;
            }
            
            // Search for paths using the spot color in all layers
            for (var layerIndex = 0; layerIndex < doc.layers.length; layerIndex++) {
                var currentLayer = doc.layers[layerIndex];
                if (currentLayer.locked || !currentLayer.visible) {
                    continue; // Skip locked or invisible layers
                }
                
                var foundItems;
                // For Spot1, search for both paths and text
                if (currentSpotColorName === "Spot1") {
                    foundItems = searchForPathsAndText(currentLayer, targetSpotColor, true);
                } else {
                    foundItems = searchForPaths(currentLayer, targetSpotColor);
                }
                pathsToMove = pathsToMove.concat(foundItems);
            }
            
            // Skip to next color if no paths found
            if (pathsToMove.length === 0) {
                continue;
            }
            
            // Get or create the target layer
            var targetLayer = getOrCreateLayer(currentSpotColorName);
            processedLayers.push(targetLayer); // Track this layer
            
            // Move all found paths to the target layer
            var movedCount = 0;
            for (var i = 0; i < pathsToMove.length; i++) {
                try {
                    pathsToMove[i].move(targetLayer, ElementPlacement.PLACEATEND);
                    movedCount++;
                } catch (e) {
                    // Continue with other items if one fails to move
                    continue;
                }
            }
            
            totalMovedPaths += movedCount;
            
            // Standardize stroke weights to 5pt for all stroked paths in this layer
            standardizeStrokeWeights(targetLayer);
            
            // Special processing for specific layers
            if (currentSpotColorName === "CutThrough-Knifecut") {
                processKnifeLayer(targetLayer, targetSpotColor);
            }
            if (currentSpotColorName === "CutThrough2-Outside") {
                processCutThrough2Layer(targetLayer, targetSpotColor);
            }
            if (currentSpotColorName === "Spot1") {
                enableOverprintFillSpot1(targetLayer);
            }
            
            // Make the target layer active for the last processed color
            if (colorIndex === spotColorNames.length - 1 || 
                (colorIndex < spotColorNames.length - 1 && 
                 !findSpotColor(spotColorNames[colorIndex + 1]))) {
                doc.activeLayer = targetLayer;
            }
        }
        
        // Check appearance consistency in all processed layers
        var inconsistentLayers = [];
        var duplicateFound = false;
        
        for (var i = 0; i < processedLayers.length; i++) {
            if (!checkLayerAppearanceConsistency(processedLayers[i])) {
                inconsistentLayers.push(processedLayers[i].name);
            }
            
            // Check for duplicate paths in newly processed layers
            if (checkForDuplicatePaths(processedLayers[i])) {
                duplicateFound = true;
            }
        }
        
        // Show duplicate warning first with warning sound
        if (duplicateFound) {
            app.beep(); // Warning sound
            alert("!!!!!! CHECK FOR DUPLICATE CUTS !!!!!!");
        }
        
        // Show appearance warning if there are inconsistent layers
        if (inconsistentLayers.length > 0) {
            alert("WARNING: The following layers contain paths with inconsistent appearances:\n\n" + 
                  inconsistentLayers.join("\n"));
        }
        
        // Play completion sound
        if (totalMovedPaths > 0) {
            app.beep();
        }
              
    } catch (error) {
        alert("An error occurred: " + error.message);
    }
})();