/*
@METADATA
{
  "name": "Qty: Sum",
  "description": "Sum Quantities from Selected Text Objects",
  "version": "1.0",
  "target": "illustrator",
  "tags": ["Qty", "sum", "utility"]
}
@END_METADATA
*/


function sumQuantitiesFromSelection() {
    // Check if there's an active document
    if (app.documents.length === 0) {
        alert("No document is open.");
        return;
    }
    
    var doc = app.activeDocument;
    var selection = doc.selection;
    
    // Check if anything is selected
    if (selection.length === 0) {
        alert("Please select some text objects first.");
        return;
    }
    
    var totalQty = 0;
    var foundQties = [];
    var processedCount = 0;
    
    // Loop through all selected objects
    for (var i = 0; i < selection.length; i++) {
        var item = selection[i];
        
        // Check if the selected item is a text frame
        if (item.typename === "TextFrame") {
            processedCount++;
            var textContent = item.contents;
            
            // Find quantities in the text using regex
            var qtyMatches = findQuantities(textContent);
            
            // Add found quantities to our arrays
            for (var j = 0; j < qtyMatches.length; j++) {
                var qty = qtyMatches[j];
                totalQty += qty;
                foundQties.push(qty);
            }
        }
    }
    
    // Display results
    displayResults(totalQty, foundQties, processedCount);
}

function findQuantities(text) {
    var quantities = [];
    
    // Regular expression to find "Qty:" followed by optional whitespace and a number
    // This handles variations like "Qty: 5", "Qty:10", "Qty: 3.5", etc.
    var qtyRegex = /Qty:\s*(\d+(?:\.\d+)?)/gi;
    var match;
    
    // Find all matches in the text
    while ((match = qtyRegex.exec(text)) !== null) {
        var number = parseFloat(match[1]);
        if (!isNaN(number)) {
            quantities.push(number);
        }
    }
    
    return quantities;
}

function displayResults(total, quantities, textFrameCount) {
    var message = "Quantity Summary:\n\n";
    message += "Text frames processed: " + textFrameCount + "\n";
    
    if (quantities.length === 0) {
        message += "No quantities found starting with 'Qty:'\n\n";
        message += "Make sure your text contains patterns like:\n";
        message += "• Qty: 5\n";
        message += "• Qty:10\n";
        message += "• Qty: 3.5";
    } else {
        message += "Quantities found: " + quantities.join(", ") + "\n";
        message += "Total sum: " + total;
        
        // If there are decimal places, show a cleaner format
        if (total % 1 !== 0) {
            message += " (" + total.toFixed(2) + ")";
        }
    }
    
    alert(message);
}

// Run the script
try {
    sumQuantitiesFromSelection();
} catch (error) {
    alert("An error occurred: " + error.message);
}