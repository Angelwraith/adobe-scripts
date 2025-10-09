/*
@METADATA
{
  "name": "Scale Selection to Target Size",
  "description": "Select a key object, enter target dimensions, and scale everything proportionally",
  "version": "1.0",
  "target": "illustrator",
  "tags": ["scale", "size", "processors"]
}
@END_METADATA
*/
function main() {
    if (app.documents.length === 0) {
        alert("Please open a document first.");
        return;
    }
    
    var doc = app.activeDocument;
    var sel = doc.selection;
    
    if (sel.length === 0) {
        alert("Please select at least one object.\n\nThe first selected object will be used as the key object for scaling.");
        return;
    }
    
    // The first selected object is the key object
    var keyObject = sel[0];
    
    // Get current dimensions of key object
    var bounds = keyObject.geometricBounds; // [left, top, right, bottom]
    var currentWidth = bounds[2] - bounds[0];
    var currentHeight = bounds[1] - bounds[3];
    
    // Convert from points to inches (72 points = 1 inch)
    var currentWidthInches = currentWidth / 72;
    var currentHeightInches = currentHeight / 72;
    
    // Create dialog
    var dialog = new Window("dialog", "Scale to Target Size");
    
    // Current size display
    dialog.add("statictext", undefined, "Current key object size:");
    dialog.add("statictext", undefined, "Width: " + currentWidthInches.toFixed(4) + '"  Height: ' + currentHeightInches.toFixed(4) + '"');
    dialog.add("statictext", undefined, ""); // Spacer
    
    // Target width input
    var widthGroup = dialog.add("group");
    widthGroup.add("statictext", undefined, "Target Width (inches):");
    var widthInput = widthGroup.add("edittext", undefined, "");
    widthInput.characters = 10;
    
    // Target height input
    var heightGroup = dialog.add("group");
    heightGroup.add("statictext", undefined, "Target Height (inches):");
    var heightInput = heightGroup.add("edittext", undefined, "");
    heightInput.characters = 10;
    
    dialog.add("statictext", undefined, ""); // Spacer
    
    // Scale mode radio buttons
    var scaleGroup = dialog.add("panel", undefined, "Scale Method");
    var radioWidth = scaleGroup.add("radiobutton", undefined, "Scale by Width (maintain aspect ratio)");
    var radioHeight = scaleGroup.add("radiobutton", undefined, "Scale by Height (maintain aspect ratio)");
    var radioBoth = scaleGroup.add("radiobutton", undefined, "Scale to exact dimensions (may distort)");
    var radioFit = scaleGroup.add("radiobutton", undefined, "Fit proportionally (fit inside dimensions)");
    
    radioWidth.value = true; // Default selection
    
    dialog.add("statictext", undefined, ""); // Spacer
    
    // Info text
    var infoText = dialog.add("statictext", undefined, "All selected objects will be scaled together.", {multiline: true});
    
    // Buttons
    var buttonGroup = dialog.add("group");
    buttonGroup.add("button", undefined, "OK", {name: "ok"});
    buttonGroup.add("button", undefined, "Cancel", {name: "cancel"});
    
    // Show dialog
    if (dialog.show() == 2) return; // User clicked Cancel
    
    // Get target dimensions and validate based on scale method
    var targetWidth = parseFloat(widthInput.text);
    var targetHeight = parseFloat(heightInput.text);
    
    // Validate inputs based on selected scale method
    if (radioWidth.value) {
        // Only need width
        if (isNaN(targetWidth) || targetWidth <= 0) {
            alert("Please enter a valid positive number for Target Width.");
            return;
        }
    } else if (radioHeight.value) {
        // Only need height
        if (isNaN(targetHeight) || targetHeight <= 0) {
            alert("Please enter a valid positive number for Target Height.");
            return;
        }
    } else {
        // Need both dimensions for "exact" or "fit" modes
        if (isNaN(targetWidth) || targetWidth <= 0 || isNaN(targetHeight) || targetHeight <= 0) {
            alert("Please enter valid positive numbers for both dimensions.");
            return;
        }
    }
    
    // Convert target dimensions to points
    var targetWidthPts = targetWidth * 72;
    var targetHeightPts = targetHeight * 72;
    
    // Calculate scale factors
    var scaleX, scaleY;
    
    if (radioWidth.value) {
        // Scale by width, maintain aspect ratio
        scaleX = scaleY = (targetWidthPts / currentWidth) * 100;
    } else if (radioHeight.value) {
        // Scale by height, maintain aspect ratio
        scaleX = scaleY = (targetHeightPts / currentHeight) * 100;
    } else if (radioBoth.value) {
        // Scale to exact dimensions (may distort)
        scaleX = (targetWidthPts / currentWidth) * 100;
        scaleY = (targetHeightPts / currentHeight) * 100;
    } else if (radioFit.value) {
        // Fit proportionally - use whichever scale is smaller
        var scaleByWidth = targetWidthPts / currentWidth;
        var scaleByHeight = targetHeightPts / currentHeight;
        scaleX = scaleY = Math.min(scaleByWidth, scaleByHeight) * 100;
    }
    
    // Get the center point of the entire selection for scaling
    var selBounds = getSelectionBounds(sel);
    var centerX = (selBounds[0] + selBounds[2]) / 2;
    var centerY = (selBounds[1] + selBounds[3]) / 2;
    
    // Group all selected objects temporarily to maintain relative positioning
    var wasGrouped = false;
    var tempGroup;
    
    if (sel.length > 1) {
        // Check if selection is already a single group
        if (sel.length === 1 && sel[0].typename === "GroupItem") {
            tempGroup = sel[0];
            wasGrouped = true;
        } else {
            // Create temporary group
            tempGroup = doc.groupItems.add();
            for (var i = sel.length - 1; i >= 0; i--) {
                sel[i].moveToBeginning(tempGroup);
            }
        }
        
        // Scale the group as a whole
        tempGroup.resize(scaleX, scaleY, true, true, true, true, scaleX, Transformation.CENTER);
        
        // Ungroup if we created a temporary group
        if (!wasGrouped) {
            var groupItems = [];
            for (var i = 0; i < tempGroup.pageItems.length; i++) {
                groupItems.push(tempGroup.pageItems[i]);
            }
            for (var i = groupItems.length - 1; i >= 0; i--) {
                groupItems[i].moveToBeginning(tempGroup.parent);
            }
            tempGroup.remove();
        }
    } else {
        // Single object, scale normally
        sel[0].resize(scaleX, scaleY, true, true, true, true, scaleX, Transformation.CENTER);
    }
    
    // Report results
    var finalBounds = keyObject.geometricBounds;
    var finalWidth = (finalBounds[2] - finalBounds[0]) / 72;
    var finalHeight = (finalBounds[1] - finalBounds[3]) / 72;
    
    alert("Scaling complete!\n\n" +
          "Key object scaled to:\n" +
          "Width: " + finalWidth.toFixed(4) + '"\n' +
          "Height: " + finalHeight.toFixed(4) + '"\n\n' +
          "Scale factors applied:\n" +
          "X: " + scaleX.toFixed(2) + "%\n" +
          "Y: " + scaleY.toFixed(2) + "%");
}

// Helper function to get bounds of entire selection
function getSelectionBounds(selection) {
    var left = Infinity, top = -Infinity, right = -Infinity, bottom = Infinity;
    
    for (var i = 0; i < selection.length; i++) {
        var bounds = selection[i].geometricBounds;
        left = Math.min(left, bounds[0]);
        top = Math.max(top, bounds[1]);
        right = Math.max(right, bounds[2]);
        bottom = Math.min(bottom, bounds[3]);
    }
    
    return [left, top, right, bottom];
}

// Run the script
main();