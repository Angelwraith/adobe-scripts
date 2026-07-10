/*
@METADATA
{
  "name": "Cut Path Separator",
  "description": "Organize Cut/Print Data Into Respective Layers",
  "version": "1.5",
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
    
    // Check for Aluminum material safety warning by checking artboard names
    var hasAluminum = false;
    for (var i = 0; i < doc.artboards.length; i++) {
        if (doc.artboards[i].name.indexOf("Alum") !== -1) {
            hasAluminum = true;
            break;
        }
    }
    
    if (hasAluminum) {
        alert("For router safety, Aluminum signs must be a minimum of 2\" from cut to cut.");
    }
    
    var spotColorNames = ["CutThrough2-Outside", "CutThrough1-Inside", "PinMountHoles-Inside", "CutThrough-Knifecut", "CutContour", "Spot1"];
    var totalMovedPaths = 0;
    var targetStrokeWidth = 5 / doc.scaleFactor; // Adjust for Large Canvas documents
    var REG_MARK_NAME = "RegMark"; // Name PRIME Flex gives each registration dot group
    var REG_DOTS_PER_ARTBOARD = 5; // Exactly this many reg dots required on any artboard that uses them
    
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
                var compoundFound = false;
                
                // Check the compound path item's own stroke/fill directly
                try {
                    if (item.strokeColor && item.strokeColor.typename === "SpotColor") {
                        if (item.strokeColor.spot.name === spotColor.name) {
                            compoundFound = true;
                        }
                    }
                } catch(e) {}
                
                try {
                    if (!compoundFound && item.fillColor && item.fillColor.typename === "SpotColor") {
                        if (item.fillColor.spot.name === spotColor.name) {
                            compoundFound = true;
                        }
                    }
                } catch(e) {}
                
                // Also check child paths as fallback
                if (!compoundFound) {
                    try {
                        for (var j = 0; j < item.pathItems.length; j++) {
                            if (usesSpotColor(item.pathItems[j], spotColor)) {
                                compoundFound = true;
                                break;
                            }
                        }
                    } catch(e) {}
                }
                
                if (compoundFound) {
                    foundPaths.push(item);
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
    
    // Validate the REG layer before doing anything else.
    // Rules: (1) only PRIME Flex reg marks (groups named REG_MARK_NAME) may live on the REG
    // layer; (2) any artboard that has reg marks must have exactly REG_DOTS_PER_ARTBOARD of
    // them (artboards with zero are fine - not all materials are cut on the flatbed cutter).
    // On any violation, show a warning listing the specifics and stop (no beep).
    function getRegLayer() {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === "REG") {
                return doc.layers[i];
            }
        }
        return null;
    }

    function describeItem(item) {
        var label = item.name && item.name !== "" ? '"' + item.name + '"' : "(unnamed)";
        return item.typename + " " + label;
    }

    function itemCenter(item) {
        // geometricBounds = [left, top, right, bottom]
        var b = item.geometricBounds;
        return { x: (b[0] + b[2]) / 2, y: (b[1] + b[3]) / 2 };
    }

    function centerInArtboard(center, rect) {
        // artboardRect = [left, top, right, bottom]
        return center.x >= rect[0] && center.x <= rect[2] &&
               center.y <= rect[1] && center.y >= rect[3];
    }

    // Structural fingerprint of a PRIME Flex reg dot, used so legacy files whose dots
    // predate the "RegMark" name still validate. A reg dot is a CLIPPED group holding
    // exactly two concentric circles: a larger unfilled outer (the clip boundary) and a
    // smaller filled inner, with an inner:outer diameter ratio of ~0.25 (0.125" / 0.5").
    // Using the ratio (not absolute sizes) keeps this scale-independent for Large Canvas.
    function isCircle(p) {
        if (!p || p.typename !== "PathItem" || !p.closed) return false;
        var b = p.geometricBounds; // [left, top, right, bottom]
        var w = b[2] - b[0];
        var h = b[1] - b[3];
        if (w <= 0 || h <= 0) return false;
        return Math.abs(w - h) / Math.max(w, h) < 0.05; // near-square bbox => round
    }

    function avgDiameter(p) {
        var b = p.geometricBounds;
        return ((b[2] - b[0]) + (b[1] - b[3])) / 2;
    }

    function isRegDotShape(item) {
        if (!item || item.typename !== "GroupItem" || !item.clipped) return false;
        if (item.pageItems.length !== 2 || item.pathItems.length !== 2) return false;
        var p0 = item.pathItems[0];
        var p1 = item.pathItems[1];
        if (!isCircle(p0) || !isCircle(p1)) return false;
        var d0 = avgDiameter(p0);
        var d1 = avgDiameter(p1);
        var outerD = Math.max(d0, d1);
        var innerD = Math.min(d0, d1);
        if (outerD <= 0) return false;
        var inner = (d0 <= d1) ? p0 : p1;
        var outer = (d0 <= d1) ? p1 : p0;
        // Concentric ratio ~0.25 (0.125"/0.5"), with a little tolerance for rounding
        if (Math.abs((innerD / outerD) - 0.25) > 0.08) return false;
        // Outer is the unfilled clip boundary; inner is the filled dot
        if (outer.filled) return false;
        if (!inner.filled) return false;
        return true;
    }

    // A valid reg mark is one PFT named "RegMark" OR one whose shape matches the dot
    // fingerprint above (covers files made before the naming change).
    function isValidRegMark(item) {
        if (item.typename === "GroupItem" && item.name === REG_MARK_NAME) return true;
        return isRegDotShape(item);
    }

    function validateRegLayer() {
        var regLayer = getRegLayer();
        if (!regLayer) {
            return true; // No REG layer -> nothing to validate
        }

        var strays = [];
        var regMarks = [];

        // Look only at the layer's direct children (nested dot circles have the group as parent)
        for (var i = 0; i < regLayer.pageItems.length; i++) {
            var item = regLayer.pageItems[i];
            if (item.parent !== regLayer) {
                continue;
            }
            if (isValidRegMark(item)) {
                regMarks.push(item);
            } else {
                strays.push(item);
            }
        }

        // Count reg marks per artboard by center containment
        var countViolations = [];
        for (var a = 0; a < doc.artboards.length; a++) {
            var rect = doc.artboards[a].artboardRect;
            var count = 0;
            for (var r = 0; r < regMarks.length; r++) {
                if (centerInArtboard(itemCenter(regMarks[r]), rect)) {
                    count++;
                }
            }
            if (count > 0 && count !== REG_DOTS_PER_ARTBOARD) {
                countViolations.push('   - "' + doc.artboards[a].name + '": ' + count +
                    " (expected " + REG_DOTS_PER_ARTBOARD + ")");
            }
        }

        if (strays.length === 0 && countViolations.length === 0) {
            return true;
        }

        var msg = "REG layer validation failed. Cut Path Separator has stopped.\n";
        if (strays.length > 0) {
            var strayList = [];
            for (var s = 0; s < strays.length; s++) {
                strayList.push("   - " + describeItem(strays[s]));
            }
            msg += "\nOnly PRIME Flex registration marks belong on the REG layer, but found:\n" +
                   strayList.join("\n") + "\n";
        }
        if (countViolations.length > 0) {
            msg += "\nArtboards with the wrong number of registration marks:\n" +
                   countViolations.join("\n") + "\n";
        }
        msg += "\nPlease correct the REG layer, then run again.";
        alert(msg);
        return false;
    }

    if (!validateRegLayer()) {
        return; // Stop CPS entirely on any REG layer problem
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
                if (currentLayer.locked) {
                    // Temporarily unlock to search, then re-lock after
                    currentLayer.locked = false;
                    var foundItems;
                    if (currentSpotColorName === "Spot1") {
                        foundItems = searchForPathsAndText(currentLayer, targetSpotColor, true);
                    } else {
                        foundItems = searchForPaths(currentLayer, targetSpotColor);
                    }
                    currentLayer.locked = true;
                    pathsToMove = pathsToMove.concat(foundItems);
                    continue;
                }
                if (!currentLayer.visible) {
                    continue; // Skip invisible layers only
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
            
            // Build a map of group relationships before moving
            var groupMap = {}; // Maps original group UUID to new group in target layer
            var itemGroupInfo = []; // Stores each item with its group ancestry
            
            // Analyze group structure for each item
            for (var i = 0; i < pathsToMove.length; i++) {
                var item = pathsToMove[i];
                var groupChain = [];
                var parent = item.parent;
                
                // Walk up the parent chain to find all ancestor groups
                while (parent && parent.typename === "GroupItem") {
                    groupChain.unshift(parent); // Add to beginning to maintain hierarchy
                    parent = parent.parent;
                }
                
                itemGroupInfo.push({
                    item: item,
                    groupChain: groupChain
                });
            }
            
            // Move all found paths to the target layer, preserving group structures
            var movedCount = 0;
            
            for (var i = 0; i < itemGroupInfo.length; i++) {
                try {
                    var info = itemGroupInfo[i];
                    var item = info.item;
                    var groupChain = info.groupChain;
                    
                    if (groupChain.length === 0) {
                        // Item is not in any group, move directly to layer
                        item.move(targetLayer, ElementPlacement.PLACEATEND);
                    } else {
                        // Item is in one or more groups - recreate the group hierarchy
                        var currentContainer = targetLayer;
                        
                        // Process each group in the chain from outermost to innermost
                        for (var j = 0; j < groupChain.length; j++) {
                            var originalGroup = groupChain[j];
                            
                            // Create a unique identifier for this group
                            var groupUUID = originalGroup.uuid;
                            
                            // Check if we've already created a matching group
                            if (!groupMap[groupUUID]) {
                                // Create new group in the current container
                                groupMap[groupUUID] = currentContainer.groupItems.add();
                            }
                            
                            // Move down to this group for the next iteration
                            currentContainer = groupMap[groupUUID];
                        }
                        
                        // Move the item to its final destination (innermost group)
                        item.move(currentContainer, ElementPlacement.PLACEATEND);
                    }
                    
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