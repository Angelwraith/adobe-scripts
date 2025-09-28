/*
@METADATA
{
  "name": "Measurement Converter",
  "description": "U.S. Customary to Decimal Converter",
  "version": "1.0",
  "target": "illustrator",
  "tags": ["measurement", "conversion", "decimal", "imperial"]
}
@END_METADATA
*/

#target illustrator

(function() {
    // Enhanced cleanup - main context
    try {
        if (typeof measurementPalette !== 'undefined' && measurementPalette instanceof Window) {
            measurementPalette.close();
        }
    } catch (e) {
        // Ignore cleanup errors
    }
    
    // Create palette window
    var measurementPalette = new Window("palette", "Measurement Converter");
    measurementPalette.orientation = "column";
    measurementPalette.alignChildren = "left";
    measurementPalette.preferredSize.width = 320;
    measurementPalette.spacing = 10;
    measurementPalette.margins = 15;
    
    // Title
    var title = measurementPalette.add("statictext", undefined, "U.S. Customary â†' Decimal");
    title.graphics.font = ScriptUI.newFont("Arial", "BOLD", 12);
    
    // Input section
    var inputGroup = measurementPalette.add("group");
    inputGroup.orientation = "column";
    inputGroup.alignChildren = "left";
    inputGroup.spacing = 5;
    
    inputGroup.add("statictext", undefined, "Enter measurement or select text:");
    var inputField = inputGroup.add("edittext", undefined, "");
    inputField.preferredSize.width = 290;
    inputField.preferredSize.height = 22;
    
    // Output section
    var outputGroup = measurementPalette.add("group");
    outputGroup.orientation = "column";
    outputGroup.alignChildren = "left";
    outputGroup.spacing = 5;
    
    outputGroup.add("statictext", undefined, "Decimal result:");
    var outputField = outputGroup.add("edittext", undefined, "");
    outputField.preferredSize.width = 290;
    outputField.preferredSize.height = 22;
    
    // Buttons
    var buttonGroup = measurementPalette.add("group");
    buttonGroup.alignment = "center";
    buttonGroup.spacing = 10;
    
    var convertButton = buttonGroup.add("button", undefined, "Convert");
    var resetButton = buttonGroup.add("button", undefined, "Reset");
    var closeButton = buttonGroup.add("button", undefined, "Close");
    
    // Get selected text from document - enhanced for different contexts
    function getSelectedText() {
        try {
            // Extra safety checks for different execution contexts
            if (!app || typeof app.documents === 'undefined') {
                return null;
            }
            
            if (app.documents.length === 0) {
                return null;
            }
            
            var doc = app.activeDocument;
            if (!doc || !doc.selection) {
                return null;
            }
            
            var sel = doc.selection;
            if (sel.length === 0) {
                return null;
            }
            
            for (var i = 0; i < sel.length; i++) {
                if (sel[i] && sel[i].typename === "TextFrame") {
                    var textContent = sel[i].contents;
                    if (textContent && textContent !== "") {
                        return textContent;
                    }
                }
            }
            return null;
        } catch (e) {
            // Return null on any error instead of breaking
            return null;
        }
    }
    
    // Simple conversion function
    function convertToDecimal(input) {
        if (!input || input === "") {
            return { success: false, error: "Please enter a measurement" };
        }
        
        var totalInches = 0;
        var cleanInput = input;
        
        // Normalize quotes
        cleanInput = cleanInput.replace(/\u2019/g, "'"); // Curly apostrophe
        cleanInput = cleanInput.replace(/\u201D/g, '"'); // Curly closing quote
        cleanInput = cleanInput.replace(/\u2018/g, "'"); // Curly opening apostrophe
        cleanInput = cleanInput.replace(/\u201C/g, '"'); // Curly opening quote
        
        var foundValidPart = false;
        
        // Look for feet pattern: number followed by apostrophe
        var feetMatches = cleanInput.match(/(\d+(?:\.\d+)?)\s*'/g);
        if (feetMatches) {
            for (var i = 0; i < feetMatches.length; i++) {
                var feetStr = feetMatches[i].replace(/\s*'/g, '');
                var feet = parseFloat(feetStr);
                if (!isNaN(feet)) {
                    totalInches += feet * 12;
                    foundValidPart = true;
                }
            }
        }
        
        // Remove feet part to avoid conflicts, then look for inches
        var inchPart = cleanInput.replace(/\d+(?:\.\d+)?\s*'/g, '').replace(/^\s*/, '');
        
        // Look for mixed number pattern: "6 1/2""
        var mixedMatch = inchPart.match(/(\d+)\s+(\d+)\s*\/\s*(\d+)\s*"/);
        if (mixedMatch) {
            var wholeInches = parseFloat(mixedMatch[1]);
            var numerator = parseFloat(mixedMatch[2]);
            var denominator = parseFloat(mixedMatch[3]);
            
            if (!isNaN(wholeInches) && !isNaN(numerator) && !isNaN(denominator) && denominator !== 0) {
                var mixedValue = wholeInches + (numerator / denominator);
                totalInches += mixedValue;
                foundValidPart = true;
            }
        }
        // Look for simple inches if no mixed number found
        else {
            var simpleInchMatch = inchPart.match(/(\d+(?:\.\d+)?)\s*"/);
            if (simpleInchMatch) {
                var inches = parseFloat(simpleInchMatch[1]);
                if (!isNaN(inches)) {
                    totalInches += inches;
                    foundValidPart = true;
                }
            }
        }
        
        // Look for standalone fractions (like 1/2" by itself)
        var fractionMatch = inchPart.match(/^(\d+)\s*\/\s*(\d+)\s*"?$/);
        if (fractionMatch) {
            var num = parseFloat(fractionMatch[1]);
            var den = parseFloat(fractionMatch[2]);
            if (!isNaN(num) && !isNaN(den) && den !== 0) {
                var fractionValue = num / den;
                totalInches += fractionValue;
                foundValidPart = true;
            }
        }
        
        if (!foundValidPart) {
            return { success: false, error: "Could not parse: " + input };
        }
        
        var result = Math.round(totalInches * 10000) / 10000;
        
        return {
            success: true,
            value: result,
            formatted: result.toString() + '"'
        };
    }
// Convert button handler - using BridgeTalk for document access
convertButton.onClick = function() {
    // 1. Read the input from the text field FIRST.
    var manualInput = inputField.text;

    // 2. Clear the output field immediately (synchronously).
    // This is the correct way to prevent the duplication bug.
    outputField.text = "";

    var bt_read = new BridgeTalk();
    bt_read.target = "illustrator";
    var scriptToRunInIllustrator = "var selectedText = null; try { if (app.documents.length > 0 && app.activeDocument.selection.length > 0) { var sel = app.activeDocument.selection; for (var i = 0; i < sel.length; i++) { if (sel[i].typename === 'TextFrame') { selectedText = sel[i].contents; break; } } } } catch(e) {}; selectedText;";
    bt_read.body = scriptToRunInIllustrator;

    bt_read.onResult = function(resultObj) {
        var selectedText = resultObj.body;
        var input = "";
        var usedSelection = false;

        if (selectedText && selectedText !== "null") {
            input = selectedText;
            inputField.text = input;
            usedSelection = true;
        } else {
            // Fallback to the manually entered text we saved.
            input = manualInput;
        }

        if (!input || input.replace(/^\s+|\s+$/g, '') === "") {
            outputField.text = "Error: Please enter a measurement";
            return;
        }

        var result = convertToDecimal(input);

        if (result.success) {
            // 3. THE FIX: Assign ONLY to the .text property.
            // Setting .textselection in an async callback is buggy and causes the duplication.
            outputField.text = result.formatted;
            outputField.active = true; // We can still focus the field.

            // If input came from selected text, update the text object
            if (usedSelection) {
                var newContent = result.formatted;
                var contentWithoutQuote = newContent.replace(/"/g, '');
                var escapedContent = contentWithoutQuote.replace(/'/g, "\\'");

                var scriptToWrite = "try { if (app.documents.length > 0 && app.activeDocument.selection[0].typename === 'TextFrame') { app.activeDocument.selection[0].contents = '" + escapedContent + "\"'; } } catch(e) {}";

                var bt_write = new BridgeTalk();
                bt_write.target = "illustrator";
                bt_write.body = scriptToWrite;
                bt_write.send();
            }
        } else {
            outputField.text = "Error: " + result.error;
        }
    }

    bt_read.send();
};
    // Reset button handler
    resetButton.onClick = function() {
        inputField.text = "";
        outputField.text = "";
        inputField.active = true;
    };
    
    // Close button handler - simple cleanup
    closeButton.onClick = function() {
        measurementPalette.close();
    };
    
    // Enter key triggers conversion
    measurementPalette.addEventListener("keydown", function(event) {
        if (event.keyName === "Enter") {
            convertButton.notify();
        }
    });
    
    // Clear output when input changes
    inputField.onChanging = function() {
        outputField.text = "";
    };
    
    // Show the palette
    measurementPalette.show();
    this.measurementPalette = measurementPalette;
    
})();