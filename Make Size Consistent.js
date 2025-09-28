// Adobe Illustrator Script: Resize All Selected Objects to Same Size
// This script resizes all selected objects to the same width and height

#target illustrator

function resizeSelectedObjects() {
    // Check if there's an active document
    if (app.documents.length === 0) {
        alert("Please open a document first.");
        return;
    }
    
    var doc = app.activeDocument;
    var selection = doc.selection;
    
    // Check if anything is selected
    if (selection.length === 0) {
        alert("Please select at least one object.");
        return;
    }
    
    // Prompt user for target width
    var targetWidth = prompt("Enter target width (in inches):", "1");
    if (targetWidth === null) return; // User cancelled
    
    // Convert to points (Illustrator's internal unit: 1 inch = 72 points)
    targetWidth = parseFloat(targetWidth) * 72;
    
    // Validate input
    if (isNaN(targetWidth) || targetWidth <= 0) {
        alert("Please enter a valid positive number for width.");
        return;
    }
    
    // Process each selected object
    var processedCount = 0;
    
    for (var i = 0; i < selection.length; i++) {
        var obj = selection[i];
        
        try {
            // Get current bounds
            var bounds = obj.visibleBounds;
            var currentWidth = bounds[2] - bounds[0];  // right - left
            var currentHeight = bounds[1] - bounds[3]; // top - bottom
            
            // Skip if current dimensions are zero
            if (currentWidth <= 0 || currentHeight <= 0) {
                continue;
            }
            
            // Calculate scale factor based on width only (proportional scaling)
            var scaleFactor = (targetWidth / currentWidth) * 100;  // Convert to percentage
            
            // Apply transformation (same scale for both width and height to maintain proportions)
            obj.resize(scaleFactor, scaleFactor);
            
            processedCount++;
            
        } catch (e) {
            // Skip objects that can't be resized (like some text objects or groups)
            continue;
        }
    }
    
    // Show completion message
    if (processedCount > 0) {
        alert("Successfully resized " + processedCount + " object(s) to " + (targetWidth/72) + "\" wide (proportionally scaled).");
    } else {
        alert("No objects could be resized. Make sure you have valid objects selected.");
    }
}

// Run the script
resizeSelectedObjects();