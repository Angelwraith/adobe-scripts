#target illustrator

/*@METADATA{
  "name": "Scale and Repeat",
  "description": "Scale artwork by ratio and repeat with spacing",
  "version": "1.0",
  "target": "illustrator",
  "tags": ["scale", "repeat", "production"]
}@END_METADATA*/

// Check if there is an active document
if (app.documents.length === 0) {
    alert("Please open a document first.");
} else {
    var doc = app.activeDocument;
    
    // Check if something is selected
    if (doc.selection.length === 0) {
        alert("Please select artwork first.");
    } else {
        // Detect Large Canvas mode
        var scaleFactor = 1;
        try {
            scaleFactor = doc.scaleFactor || 1;
        } catch (e) {
            scaleFactor = 1;
        }
        
        // Show dialog
        var dialog = new Window("dialog", "Scale and Repeat");
        dialog.alignChildren = "center";
        dialog.margins = 20;
        
        // Scale preset options
        var scaleOptions = ["1:1", "1:2", "1:4", "1:5", "1:8", "1:10", "1:16", "1:20", "1:40", "1:50", "1:100", "2:1", "4:1", "8:1", "10:1", "100:2", "Custom"];
        
        // Labels row
        var labelRow = dialog.add("group");
        labelRow.alignment = "center";
        
        var currentLabel = labelRow.add("statictext", undefined, "Current");
        currentLabel.preferredSize.width = 120;
        currentLabel.justify = "center";
        
        var targetLabel = labelRow.add("statictext", undefined, "Target");
        targetLabel.preferredSize.width = 120;
        targetLabel.justify = "center";
        
        var qtyLabel = labelRow.add("statictext", undefined, "Qty");
        qtyLabel.preferredSize.width = 120;
        qtyLabel.justify = "center";
        
        var spacingLabel = labelRow.add("statictext", undefined, "Spacing");
        spacingLabel.preferredSize.width = 120;
        spacingLabel.justify = "center";
        
        // Main row with dropdowns and separators
        var mainRow = dialog.add("group");
        mainRow.alignment = "center";
        
        var currentDropdown = mainRow.add("dropdownlist", undefined, scaleOptions);
        currentDropdown.preferredSize.width = 120;
        currentDropdown.selection = 5; // Default to 1:10
        
        var separator1 = mainRow.add("statictext", undefined, "|");
        separator1.preferredSize.width = 20;
        separator1.justify = "center";
        
        var targetDropdown = mainRow.add("dropdownlist", undefined, scaleOptions);
        targetDropdown.preferredSize.width = 120;
        targetDropdown.selection = 0; // Default to 1:1
        
        var separator2 = mainRow.add("statictext", undefined, "x");
        separator2.preferredSize.width = 20;
        separator2.justify = "center";
        
        var qtyOptions = [];
        for (var i = 1; i <= 20; i++) {
            qtyOptions.push(i.toString());
        }
        qtyOptions.push("Custom");
        
        var qtyDropdown = mainRow.add("dropdownlist", undefined, qtyOptions);
        qtyDropdown.preferredSize.width = 120;
        qtyDropdown.selection = 0; // Default to 1
        
        var separator3 = mainRow.add("statictext", undefined, "|");
        separator3.preferredSize.width = 20;
        separator3.justify = "center";
        
        var spacingOptions = ["No Spacing", "1\" Spacing"];
        var spacingDropdown = mainRow.add("dropdownlist", undefined, spacingOptions);
        spacingDropdown.preferredSize.width = 120;
        spacingDropdown.selection = 0; // Default to No Spacing
        
        // Custom input row (hidden by default, shows fields as needed)
        var customRow = dialog.add("group");
        customRow.alignment = "center";
        customRow.visible = false;
        
        // Current custom fields
        var currentNumField = customRow.add("edittext", undefined, "1");
        currentNumField.characters = 5;
        currentNumField.visible = false;
        
        var currentColon = customRow.add("statictext", undefined, ":");
        currentColon.visible = false;
        
        var currentDenomField = customRow.add("edittext", undefined, "10");
        currentDenomField.characters = 5;
        currentDenomField.visible = false;
        
        var customSeparator1 = customRow.add("statictext", undefined, "|");
        customSeparator1.preferredSize.width = 20;
        customSeparator1.justify = "center";
        customSeparator1.visible = false;
        
        // Target custom fields
        var targetNumField = customRow.add("edittext", undefined, "1");
        targetNumField.characters = 5;
        targetNumField.visible = false;
        
        var targetColon = customRow.add("statictext", undefined, ":");
        targetColon.visible = false;
        
        var targetDenomField = customRow.add("edittext", undefined, "1");
        targetDenomField.characters = 5;
        targetDenomField.visible = false;
        
        var customSeparator2 = customRow.add("statictext", undefined, "x");
        customSeparator2.preferredSize.width = 20;
        customSeparator2.justify = "center";
        customSeparator2.visible = false;
        
        // Qty custom field
        var qtyCustomField = customRow.add("edittext", undefined, "1");
        qtyCustomField.characters = 5;
        qtyCustomField.visible = false;
        
        // Event handlers for dropdowns
        var updateCustomRow = function() {
            var currentIsCustom = (currentDropdown.selection.text === "Custom");
            var targetIsCustom = (targetDropdown.selection.text === "Custom");
            var qtyIsCustom = (qtyDropdown.selection.text === "Custom");
            
            // Show/hide current custom fields
            currentNumField.visible = currentIsCustom;
            currentColon.visible = currentIsCustom;
            currentDenomField.visible = currentIsCustom;
            customSeparator1.visible = currentIsCustom || targetIsCustom || qtyIsCustom;
            
            // Show/hide target custom fields
            targetNumField.visible = targetIsCustom;
            targetColon.visible = targetIsCustom;
            targetDenomField.visible = targetIsCustom;
            customSeparator2.visible = (currentIsCustom || targetIsCustom) && qtyIsCustom;
            
            // Show/hide qty custom field
            qtyCustomField.visible = qtyIsCustom;
            
            // Show/hide entire custom row
            customRow.visible = currentIsCustom || targetIsCustom || qtyIsCustom;
            
            dialog.layout.layout(true);
        };
        
        currentDropdown.onChange = updateCustomRow;
        targetDropdown.onChange = updateCustomRow;
        qtyDropdown.onChange = updateCustomRow;
        
        // Add buttons
        var buttonGroup = dialog.add("group");
        buttonGroup.alignment = "center";
        buttonGroup.spacing = 10;
        var okButton = buttonGroup.add("button", undefined, "OK");
        var cancelButton = buttonGroup.add("button", undefined, "Cancel");
        
        // Show dialog and get result
        if (dialog.show() === 1) {
            var currentNum, currentDenom, targetNum, targetDenom, qty;
            
            // Parse current scale
            if (currentDropdown.selection.text === "Custom") {
                currentNum = parseFloat(currentNumField.text);
                currentDenom = parseFloat(currentDenomField.text);
            } else {
                var currentParts = currentDropdown.selection.text.split(":");
                currentNum = parseFloat(currentParts[0]);
                currentDenom = parseFloat(currentParts[1]);
            }
            
            // Parse target scale
            if (targetDropdown.selection.text === "Custom") {
                targetNum = parseFloat(targetNumField.text);
                targetDenom = parseFloat(targetDenomField.text);
            } else {
                var targetParts = targetDropdown.selection.text.split(":");
                targetNum = parseFloat(targetParts[0]);
                targetDenom = parseFloat(targetParts[1]);
            }
            
            // Parse quantity
            if (qtyDropdown.selection.text === "Custom") {
                qty = parseInt(qtyCustomField.text);
            } else {
                qty = parseInt(qtyDropdown.selection.text);
            }
            
            // Validate inputs
            if (isNaN(currentNum) || isNaN(currentDenom) || isNaN(targetNum) || isNaN(targetDenom) || isNaN(qty)) {
                alert("Please enter valid numbers in all fields.");
            } else if (currentDenom === 0 || targetDenom === 0) {
                alert("Denominators cannot be zero.");
            } else if (qty < 1) {
                alert("Quantity must be at least 1.");
            } else {
                // Calculate scale factor
                var currentScale = currentNum / currentDenom;
                var targetScale = targetNum / targetDenom;
                var scaleFactorMultiplier = targetScale / currentScale;
                var scalePercent = scaleFactorMultiplier * 100;
                
                // Get selected items
                var selectedItems = [];
                for (var i = 0; i < doc.selection.length; i++) {
                    selectedItems.push(doc.selection[i]);
                }
                
                // Group selection if multiple items
                var itemToProcess;
                if (selectedItems.length > 1) {
                    doc.selection = selectedItems;
                    itemToProcess = doc.groupItems.add();
                    for (var i = selectedItems.length - 1; i >= 0; i--) {
                        selectedItems[i].moveToBeginning(itemToProcess);
                    }
                } else {
                    itemToProcess = selectedItems[0];
                }
                
                // Get original bounds
                var originalBounds = itemToProcess.geometricBounds;
                var originalWidth = originalBounds[2] - originalBounds[0];
                
                // Scale the original item
                itemToProcess.resize(scalePercent, scalePercent);
                
                // Get new bounds after scaling
                var scaledBounds = itemToProcess.geometricBounds;
                var scaledWidth = scaledBounds[2] - scaledBounds[0];
                
                // Determine spacing based on user selection
                var spacingPoints = 0;
                if (spacingDropdown.selection.text === "1\" Spacing") {
                    // Spacing in points (1 inch = 72 points)
                    // For Large Canvas documents, divide by scaleFactor to get correct spacing
                    spacingPoints = 72 / scaleFactor;
                }
                
                // Create duplicates
                for (var i = 1; i < qty; i++) {
                    var duplicate = itemToProcess.duplicate();
                    var offsetDistance = (scaledWidth + spacingPoints) * i;
                    duplicate.translate(offsetDistance, 0);
                }
                
                var message = "Complete! Scaled by " + scaleFactorMultiplier.toFixed(3) + "x and created " + qty + " item(s).";
                if (scaleFactor !== 1) {
                    message += "\n(Large Canvas detected - spacing adjusted)";
                }
                alert(message);
            }
        }
    }
}
