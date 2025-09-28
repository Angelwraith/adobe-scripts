/*
@METADATA
{
  "name": "Wrap Image Namer",
  "description": "Adds Filenames Above Images for Wrap ProdProofs",
  "version": "1.0",
  "target": "illustrator",
  "tags": ["ProdProof", "namer", "wrap"]
}
@END_METADATA
*/

(function() {
    // Check if we have a document open
    if (app.documents.length === 0) {
        alert("Please open a document first.");
        return;
    }
    
    var doc = app.activeDocument;
    var selection = doc.selection;
    
    // Check if anything is selected
    if (selection.length === 0) {
        alert("Please select one or more images first.");
        return;
    }
    
    // Filter selection to only include raster images and placed items
    var images = [];
    for (var i = 0; i < selection.length; i++) {
        if (selection[i].typename === "RasterItem" || 
            selection[i].typename === "PlacedItem") {
            images.push(selection[i]);
        }
    }
    
    if (images.length === 0) {
        alert("No images found in selection. Please select raster images or placed items.");
        return;
    }
    
    // Process each image
    for (var j = 0; j < images.length; j++) {
        var image = images[j];
        
        // Get the actual filename from the linked image
        var imageName = "Image"; // fallback
        
        if (image.typename === "PlacedItem" && image.file) {
            // For placed/linked images, get the filename from the file path
            var fullPath = image.file.fsName || image.file.name;
            imageName = fullPath.split(/[\\\/]/).pop(); // Get just the filename
            // Remove file extension
            var lastDot = imageName.lastIndexOf(".");
            if (lastDot > 0) {
                imageName = imageName.substring(0, lastDot);
            }
        } else if (image.typename === "RasterItem" && image.file) {
            // For embedded raster items that still have file reference
            var fullPath = image.file.fsName || image.file.name;
            imageName = fullPath.split(/[\\\/]/).pop();
            var lastDot = imageName.lastIndexOf(".");
            if (lastDot > 0) {
                imageName = imageName.substring(0, lastDot);
            }
        } else {
            // Try to get from the object name as fallback
            imageName = image.name || "Image";
        }
        
        // Replace %20 and underscores with spaces (using simple string replacement)
        while (imageName.indexOf("%20") > -1) {
            imageName = imageName.replace("%20", " ");
        }
        while (imageName.indexOf("_") > -1) {
            imageName = imageName.replace("_", " ");
        }
        
        // Get image bounds for positioning
        var imageBounds = image.geometricBounds;
        var imageLeft = imageBounds[0];
        var imageTop = imageBounds[1];
        var imageRight = imageBounds[2];
        var imageBottom = imageBounds[3];
        
        // Calculate center X position of the image
        var centerX = imageLeft + (imageRight - imageLeft) / 2;
        
        // Calculate Y position (0.125" = 9 points above the top of the image)
        var textY = imageTop + 9; // 0.125 inch = 9 points
        
        // Create point type text frame
        var textFrame = doc.textFrames.pointText([centerX, textY]);
        
        // Set the text content
        textFrame.contents = imageName;
        
        // Format the text
        var textRange = textFrame.textRange;
        var charAttributes = textRange.characterAttributes;
        
        // Set font (Myriad Pro Black, 16pt)
        try {
            charAttributes.textFont = app.textFonts.getByName("MyriadPro-Black");
        } catch (e) {
            // If Myriad Pro Black isn't available, try alternatives
            try {
                charAttributes.textFont = app.textFonts.getByName("Myriad Pro");
                charAttributes.fontWeight = "Bold";
            } catch (e2) {
                // Use system default if Myriad Pro isn't available
                charAttributes.textFont = app.textFonts[0];
            }
        }
        
        charAttributes.size = 16;
        
        // Set paragraph alignment to center
        var paraAttributes = textRange.paragraphAttributes;
        paraAttributes.justification = Justification.CENTER;
        
        // Position the text frame so it's centered above the image
        // Since we created point type at the center X, we need to adjust for centering
        var textBounds = textFrame.geometricBounds;
        var textWidth = textBounds[2] - textBounds[0];
        textFrame.left = centerX - (textWidth / 2);
    }
    
    // Deselect all and show completion message
    doc.selection = null;
    alert("Successfully added point type above " + images.length + " image(s).");
    
})();