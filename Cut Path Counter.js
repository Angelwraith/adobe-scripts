/*@METADATA{
  "name": "Cut Path Counter",
  "description": "Count CutContour and CutThrough2-Outside paths in current selection",
  "version": "1.0",
  "target": "illustrator",
  "tags": ["selection", "cutcontour", "cutthrough", "count"]
}@END_METADATA*/

// Adobe Illustrator Script: Count Cut Paths in Selection
// This script counts CutContour and CutThrough2-Outside paths in the current selection

if (app.documents.length == 0) {
    alert("Please open a document first.");
} else if (app.selection.length == 0) {
    alert("Please select some objects first.");
} else {
    var doc = app.activeDocument;
    var selectedItems = app.selection;
    
    // Function to convert points to inches (accounting for scale factor)
    function pointsToInches(points) {
        try {
            var scaleFactor = app.activeDocument.scaleFactor;
            // Handle undefined or invalid scale factor
            if (!scaleFactor || scaleFactor <= 0 || isNaN(scaleFactor)) {
                scaleFactor = 1;
            }
            return Math.round(((points / 72) * scaleFactor) * 100) / 100; // Round to 2 decimal places
        } catch (e) {
            // Fallback to no scale factor if there's an error
            return Math.round((points / 72) * 100) / 100;
        }
    }
    
    // Function to get path dimensions (accounting for rotation)
    function getPathDimensions(pathItem) {
        var bounds = pathItem.geometricBounds;
        var width = pointsToInches(bounds[2] - bounds[0]);
        var height = pointsToInches(bounds[1] - bounds[3]);
        
        // Return dimensions in consistent format (smaller dimension first)
        if (width <= height) {
            return width + '"x' + height + '"';
        } else {
            return height + '"x' + width + '"';
        }
    }
    
    // Function to check if an item has CutContour or CutThrough2-Outside spot color
    function hasCutColor(item) {
        try {
            var colorName = null;
            
            // Check fill color
            if (item.filled && item.fillColor) {
                if (item.fillColor.typename == "SpotColor" && item.fillColor.spot) {
                    var spotName = item.fillColor.spot.name;
                    if (spotName == "CutContour" || spotName == "CutThrough2-Outside") {
                        colorName = spotName;
                    }
                }
            }
            
            // Check stroke color (if no fill color found)
            if (!colorName && item.stroked && item.strokeColor) {
                if (item.strokeColor.typename == "SpotColor" && item.strokeColor.spot) {
                    var spotName = item.strokeColor.spot.name;
                    if (spotName == "CutContour" || spotName == "CutThrough2-Outside") {
                        colorName = spotName;
                    }
                }
            }
            
            return colorName;
        } catch (e) {
            // Ignore errors for items without color properties
        }
        return null;
    }
    
    // Function to recursively process selected items
    function processSelectedItems(items, cutContourSizes, cutThroughSizes, processedItems) {
        if (!processedItems) {
            processedItems = [];
        }
        
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            
            // Skip if we've already processed this item (avoid double counting)
            var itemFound = false;
            for (var j = 0; j < processedItems.length; j++) {
                if (processedItems[j] === item) {
                    itemFound = true;
                    break;
                }
            }
            if (itemFound) continue;
            
            // Check for cut colors
            if (item.typename == "PathItem") {
                var cutColorName = hasCutColor(item);
                if (cutColorName) {
                    var size = getPathDimensions(item);
                    
                    if (cutColorName == "CutContour") {
                        if (cutContourSizes[size]) {
                            cutContourSizes[size]++;
                        } else {
                            cutContourSizes[size] = 1;
                        }
                    } else if (cutColorName == "CutThrough2-Outside") {
                        if (cutThroughSizes[size]) {
                            cutThroughSizes[size]++;
                        } else {
                            cutThroughSizes[size] = 1;
                        }
                    }
                    
                    processedItems.push(item);
                }
            }
            
            // Recursively process grouped items
            if (item.typename == "GroupItem" && item.pageItems.length > 0) {
                processSelectedItems(item.pageItems, cutContourSizes, cutThroughSizes, processedItems);
            }
        }
        
        return processedItems;
    }
    
    try {
        var cutContourSizes = {};
        var cutThroughSizes = {};
        var totalCutContour = 0;
        var totalCutThrough = 0;
        
        // Process selected items
        processSelectedItems(selectedItems, cutContourSizes, cutThroughSizes);
        
        // Count totals
        for (var size in cutContourSizes) {
            totalCutContour += cutContourSizes[size];
        }
        for (var size in cutThroughSizes) {
            totalCutThrough += cutThroughSizes[size];
        }
        
        // Build report
        var report = [];
        report.push("=== CUT PATH COUNTER - SELECTION ===");
        report.push("");
        
        // CutContour section
        if (totalCutContour > 0) {
            report.push('CUTCONTOUR PATHS:');
            
            // Sort sizes for consistent output
            var sortedContourSizes = [];
            for (var size in cutContourSizes) {
                sortedContourSizes.push(size);
            }
            sortedContourSizes.sort();
            
            for (var i = 0; i < sortedContourSizes.length; i++) {
                var size = sortedContourSizes[i];
                var count = cutContourSizes[size];
                report.push("CutContour @ " + size + " = " + count);
            }
            report.push("Total CutContour paths: " + totalCutContour);
        } else {
            report.push("CUTCONTOUR PATHS: None found");
        }
        
        report.push("");
        
        // CutThrough2-Outside section
        if (totalCutThrough > 0) {
            report.push('CUTTHROUGH2-OUTSIDE PATHS:');
            
            // Sort sizes for consistent output
            var sortedThroughSizes = [];
            for (var size in cutThroughSizes) {
                sortedThroughSizes.push(size);
            }
            sortedThroughSizes.sort();
            
            for (var i = 0; i < sortedThroughSizes.length; i++) {
                var size = sortedThroughSizes[i];
                var count = cutThroughSizes[size];
                report.push("CutThrough2-Outside @ " + size + " = " + count);
            }
            report.push("Total CutThrough2-Outside paths: " + totalCutThrough);
        } else {
            report.push("CUTTHROUGH2-OUTSIDE PATHS: None found");
        }
        
        report.push("");
        report.push("=== SUMMARY ===");
        report.push("CutContour paths: " + totalCutContour);
        report.push("CutThrough2-Outside paths: " + totalCutThrough);
        report.push("Total cut paths: " + (totalCutContour + totalCutThrough));
        
        // Show results in popup dialog
        var reportText = report.join("\n");
        
        // Create modal dialog
        var dialog = new Window("dialog", "Cut Path Counter - Selection Results");
        dialog.orientation = "column";
        dialog.alignChildren = "fill";
        dialog.spacing = 10;
        dialog.margins = 16;
        
        var textArea = dialog.add("edittext", undefined, reportText, {multiline: true, scrolling: true, readonly: true});
        textArea.preferredSize.width = 500;
        textArea.preferredSize.height = 350;
        
        // Set monospace font
        try {
            textArea.graphics.font = ScriptUI.newFont("Monaco", "Regular", 11);
        } catch (e) {
            try {
                textArea.graphics.font = ScriptUI.newFont("Courier New", "Regular", 11);
            } catch (e2) {}
        }
        
        var buttonGroup = dialog.add("group");
        buttonGroup.alignment = "center";
        buttonGroup.spacing = 10;
        
        var closeButton = buttonGroup.add("button", undefined, "Close");
        
        closeButton.onClick = function() { dialog.close(); };
        closeButton.active = true;
        
        dialog.show();
        
    } catch (error) {
        alert("Error counting cut paths: " + error.toString());
    }
}