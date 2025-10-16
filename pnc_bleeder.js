/*@METADATA{
  "name": "PnC Bleeder",
  "description": "Automatically generates intelligent color-matched bleed/trap paths for vinyl cutting on complex multi-color artwork",
  "version": "1.0",
  "target": "illustrator",
  "tags": ["vinyl cutting", "trapping", "bleed generation"]
}@END_METADATA*/

// PnC Bleeder - Intelligent Vinyl Cutting Bleed Generator
// Creates color-matched trap paths for clean vinyl cuts

(function() {
    
    // Configuration
    var config = {
        bleedWidth: 0.02, // inches - adjust based on your cutting tolerance
        debugMode: false // set to true to see what's happening
    };
    
    // Main execution
    try {
        if (app.documents.length === 0) {
            alert("Please open a document first.");
            return;
        }
        
        var doc = app.activeDocument;
        var sel = doc.selection;
        
        if (sel.length === 0) {
            alert("Please select the artwork you want to add bleed to.");
            return;
        }
        
        // Show dialog to get bleed width
        var dialog = new Window("dialog", "PnC Bleeder Settings");
        dialog.add("statictext", undefined, "Bleed/Trap Width (inches):");
        var bleedInput = dialog.add("edittext", undefined, config.bleedWidth.toString());
        bleedInput.characters = 10;
        
        var buttonGroup = dialog.add("group");
        buttonGroup.add("button", undefined, "OK", {name: "ok"});
        buttonGroup.add("button", undefined, "Cancel", {name: "cancel"});
        
        if (dialog.show() === 2) return; // User cancelled
        
        config.bleedWidth = parseFloat(bleedInput.text);
        if (isNaN(config.bleedWidth) || config.bleedWidth <= 0) {
            alert("Please enter a valid bleed width.");
            return;
        }
        
        // Convert inches to points (Illustrator's internal unit: 72 points = 1 inch)
        var bleedPoints = config.bleedWidth * 72;
        
        // Create a new layer for bleed paths at the bottom
        var bleedLayer = doc.layers.add();
        bleedLayer.name = "Bleed";
        bleedLayer.move(doc.layers[doc.layers.length - 1], ElementPlacement.PLACEAFTER);
        
        var processedCount = 0;
        var errorCount = 0;
        var skippedCount = 0;
        var compoundCount = 0;
        var pathCount = 0;
        var skippedItems = []; // Track what was skipped
        
        // Collect all items to process
        var itemsToProcess = [];
        collectItems(sel, itemsToProcess);
        
        // Count types
        for (var i = 0; i < itemsToProcess.length; i++) {
            var itemType = itemsToProcess[i].typename;
            if (itemType === "CompoundPathItem") compoundCount++;
            if (itemType === "PathItem") pathCount++;
        }
        
        // Process each item
        for (var i = 0; i < itemsToProcess.length; i++) {
            try {
                var result = processPathItem(itemsToProcess[i], bleedLayer, bleedPoints, doc);
                if (result === "skipped") {
                    skippedCount++;
                    skippedItems.push({
                        name: itemsToProcess[i].name || "unnamed",
                        type: itemsToProcess[i].typename,
                        index: i
                    });
                } else {
                    processedCount++;
                }
            } catch (e) {
                errorCount++;
                if (config.debugMode) {
                    alert("Error processing item " + i + " (" + itemsToProcess[i].typename + "): " + e.toString());
                }
            }
        }
        
        // Build skipped items report
        var skippedReport = "";
        if (skippedItems.length > 0) {
            skippedReport = "\n\nSkipped items:";
            for (var i = 0; i < Math.min(skippedItems.length, 10); i++) {
                skippedReport += "\n  " + skippedItems[i].type + " '" + skippedItems[i].name + "'";
            }
            if (skippedItems.length > 10) {
                skippedReport += "\n  ... and " + (skippedItems.length - 10) + " more";
            }
        }
        
        // Summary
        alert("PnC Bleeder Complete!\n\n" +
              "Selected items breakdown:\n" +
              "  Paths: " + pathCount + "\n" +
              "  Compound Paths: " + compoundCount + "\n\n" +
              "Results:\n" +
              "  Processed: " + processedCount + " items\n" +
              "  Skipped (no fill): " + skippedCount + " items\n" +
              "  Errors: " + errorCount + " items" +
              skippedReport + "\n\n" +
              "Bleed paths created on layer: 'Bleed'");
        
    } catch (e) {
        alert("Error: " + e.toString());
    }
    
    // Collect all items recursively from selection
    function collectItems(items, collection) {
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (item.typename === "GroupItem") {
                collectItems(item.pageItems, collection);
            } else if (item.typename === "PathItem" || item.typename === "CompoundPathItem") {
                collection.push(item);
            }
        }
    }
    
    // Process a single path item and create bleed using pathfinder operations
    function processPathItem(item, bleedLayer, bleedPoints, doc) {
        
        // Get the fill color of the original item
        var fillColor = getItemColor(item);
        if (!fillColor) {
            if (config.debugMode) {
                alert("Skipping " + item.typename + " '" + (item.name || "unnamed") + "'\nReason: No fill color detected");
            }
            return "skipped";
        }
        
        if (config.debugMode) {
            alert("Processing " + item.typename + " '" + (item.name || "unnamed") + "'\nHas fill: Yes\nColor type: " + fillColor.typename + "\nCreating bleed...");
        }
        
        // Create only the outer duplicate for bleed
        var outerDup = item.duplicate();
        
        // Move to bleed layer
        outerDup.move(bleedLayer, ElementPlacement.PLACEATEND);
        
        if (config.debugMode) {
            alert("Duplicate created and moved.\nItem type: " + item.typename + "\nAttempting to set stroke...\nBleed width: " + (bleedPoints * 2) + " points");
        }
        
        // For compound paths, we need to set properties on the sub-paths
        if (item.typename === "CompoundPathItem") {
            if (config.debugMode) {
                alert("This is a CompoundPathItem with " + outerDup.pathItems.length + " sub-paths.\nSetting stroke on each sub-path...");
            }
            
            // Set stroke on all sub-paths
            for (var i = 0; i < outerDup.pathItems.length; i++) {
                var subPath = outerDup.pathItems[i];
                subPath.stroked = true;
                subPath.strokeWidth = bleedPoints * 2;
                subPath.strokeColor = copyColor(fillColor);
                // Keep the fill so the shape is visible
                subPath.filled = true;
                subPath.fillColor = copyColor(fillColor);
                
                if (config.debugMode && i === 0) {
                    alert("Sub-path " + i + " stroke set:\nstroked = " + subPath.stroked + "\nstrokeWidth = " + subPath.strokeWidth);
                }
            }
        } else {
            // Regular path - apply directly
            outerDup.stroked = true;
            outerDup.strokeWidth = bleedPoints * 2;
            outerDup.strokeColor = copyColor(fillColor);
            // Keep the fill so the shape is visible
            outerDup.filled = true;
            outerDup.fillColor = copyColor(fillColor);
            
            if (config.debugMode) {
                alert("Regular PathItem stroke set:\nstroked = " + outerDup.stroked + "\nstrokeWidth = " + outerDup.strokeWidth);
            }
        }
        
        // Name it for reference
        outerDup.name = "Bleed for: " + (item.name || "unnamed");
        
        if (config.debugMode) {
            alert("Bleed created successfully for " + item.typename + "\n\nCheck the bleed layer for:\n- " + outerDup.name);
        }
        
        return "processed";
    }
    
    // Get color from an item
    function getItemColor(item) {
        if (item.typename === "PathItem") {
            if (item.filled && item.fillColor) {
                return item.fillColor;
            }
        } else if (item.typename === "CompoundPathItem") {
            // Get color from first sub-path - need to check if filled property exists
            if (item.pathItems && item.pathItems.length > 0) {
                var firstPath = item.pathItems[0];
                // Check if the path has fillColor regardless of filled property
                if (firstPath.fillColor) {
                    if (config.debugMode) {
                        alert("Found fillColor on compound path sub-item\nColor type: " + firstPath.fillColor.typename);
                    }
                    return firstPath.fillColor;
                }
            }
        }
        
        if (config.debugMode) {
            alert("No fill color found for " + item.typename);
        }
        return null;
    }
    
    // Copy a color object
    function copyColor(color) {
        if (!color) return null;
        
        var newColor;
        
        if (color.typename === "RGBColor") {
            newColor = new RGBColor();
            newColor.red = color.red;
            newColor.green = color.green;
            newColor.blue = color.blue;
        } else if (color.typename === "CMYKColor") {
            newColor = new CMYKColor();
            newColor.cyan = color.cyan;
            newColor.magenta = color.magenta;
            newColor.yellow = color.yellow;
            newColor.black = color.black;
        } else if (color.typename === "GrayColor") {
            newColor = new GrayColor();
            newColor.gray = color.gray;
        } else if (color.typename === "SpotColor") {
            newColor = new SpotColor();
            newColor.spot = color.spot;
            newColor.tint = color.tint;
        } else {
            // Default to gray
            newColor = new GrayColor();
            newColor.gray = 50;
        }
        
        return newColor;
    }
    
})();