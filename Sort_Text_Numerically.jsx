/*@METADATA{
  "name": "Sort Text Numerically",
  "description": "Sorts selected text frames numerically from left to right based on numbers found in the text content",
  "version": "1.0",
  "target": "illustrator",
  "tags": ["text", "sort", "organize"]
}@END_METADATA*/

// Main function
function sortTextNumerically() {
    // Check if there's an active document
    if (app.documents.length === 0) {
        alert("Please open a document first.");
        return;
    }
    
    var doc = app.activeDocument;
    
    // Check if there are selected items
    if (doc.selection.length === 0) {
        alert("Please select text frames to sort.");
        return;
    }
    
    // Collect all text frames from selection
    var textFrames = [];
    for (var i = 0; i < doc.selection.length; i++) {
        var item = doc.selection[i];
        if (item.typename === "TextFrame") {
            textFrames.push(item);
        }
    }
    
    // Check if we found any text frames
    if (textFrames.length === 0) {
        alert("No text frames found in selection.");
        return;
    }
    
    // Create array of objects with text frame and extracted number
    var textData = [];
    for (var i = 0; i < textFrames.length; i++) {
        var frame = textFrames[i];
        var text = frame.contents;
        var number = extractNumber(text);
        
        textData.push({
            frame: frame,
            number: number,
            originalY: frame.top
        });
    }
    
    // Sort by number
    textData.sort(function(a, b) {
        return a.number - b.number;
    });
    
    // Calculate average Y position to maintain
    var totalY = 0;
    for (var i = 0; i < textData.length; i++) {
        totalY += textData[i].originalY;
    }
    var averageY = totalY / textData.length;
    
    // Calculate spacing
    var startX = 100; // Starting X position (left margin)
    var spacing = 100; // Space between items
    
    // Reposition text frames from left to right
    for (var i = 0; i < textData.length; i++) {
        var frame = textData[i].frame;
        var newX = startX + (i * spacing);
        
        // Move the frame to new position
        frame.left = newX;
        frame.top = averageY;
    }
    
    alert("Sorted " + textFrames.length + " text frames numerically from left to right.");
}

// Function to extract the first number from a text string
function extractNumber(text) {
    // Remove all non-digit characters and extract the first number
    var matches = text.match(/\d+/);
    
    if (matches && matches.length > 0) {
        return parseInt(matches[0], 10);
    }
    
    // If no number found, return a very high number to sort it at the end
    return 999999;
}

// Run the script
sortTextNumerically();
