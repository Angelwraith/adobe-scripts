#target illustrator

/*@METADATA{
  "name": "Scale and Repeat",
  "description": "Scale artwork by ratio and repeat with spacing",
  "version": "1.2",
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
        dialog.orientation = "column";
        dialog.alignChildren = "fill";
        dialog.margins = 20;
        dialog.spacing = 10;
        
        var LABEL_W = 70;
        var FIELD_W = 120;
        
        // Scale preset options
        var scaleOptions = ["1:1", "1:2", "1:4", "1:5", "1:8", "1:10", "1:16", "1:20", "1:40", "1:50", "1:100", "2:1", "4:1", "8:1", "10:1", "100:2", "Custom"];
        
        // Helper: one row with label on the left, controls on the right
        var addRow = function(labelText) {
            var row = dialog.add("group");
            row.orientation = "row";
            row.alignChildren = ["left", "center"];
            var lbl = row.add("statictext", undefined, labelText);
            lbl.preferredSize.width = LABEL_W;
            lbl.justify = "right";
            return row;
        };
        
        // Helper: ratio row (dropdown + custom num:denom fields shown when "Custom")
        var addRatioRow = function(labelText, defNum, defDenom) {
            var row = addRow(labelText);
            var dd = row.add("dropdownlist", undefined, scaleOptions);
            dd.preferredSize.width = FIELD_W;
            dd.selection = 0; // Default to 1:1
            var numField = row.add("edittext", undefined, defNum);
            numField.characters = 5;
            var colon = row.add("statictext", undefined, ":");
            var denomField = row.add("edittext", undefined, defDenom);
            denomField.characters = 5;
            var setCustomVisible = function() {
                var isCustom = (dd.selection.text === "Custom");
                numField.visible = isCustom;
                colon.visible = isCustom;
                denomField.visible = isCustom;
            };
            dd.onChange = setCustomVisible;
            setCustomVisible();
            return { dropdown: dd, num: numField, denom: denomField };
        };
        
        var currentRow = addRatioRow("Current:", "1", "10");
        var targetRow = addRatioRow("Target:", "1", "1");
        var currentDropdown = currentRow.dropdown;
        var targetDropdown = targetRow.dropdown;
        
        // Qty row: preset dropdown + always-visible entry field (entry wins if filled)
        var qtyRow = addRow("Qty:");
        var qtyOptions = [];
        for (var i = 1; i <= 20; i++) {
            qtyOptions.push(i.toString());
        }
        var qtyDropdown = qtyRow.add("dropdownlist", undefined, qtyOptions);
        qtyDropdown.preferredSize.width = FIELD_W;
        qtyDropdown.selection = 0; // Default to 1
        var qtyField = qtyRow.add("edittext", undefined, "");
        qtyField.characters = 6;
        qtyField.helpTip = "Type a quantity to override the dropdown";
        
        // Spacing row
        var spacingRow = addRow("Spacing:");
        var spacingOptions = ["No Spacing", "1\" Spacing"];
        var spacingDropdown = spacingRow.add("dropdownlist", undefined, spacingOptions);
        spacingDropdown.preferredSize.width = FIELD_W;
        spacingDropdown.selection = 0; // Default to No Spacing
        
        // Add buttons
        var buttonGroup = dialog.add("group");
        buttonGroup.alignment = "center";
        buttonGroup.spacing = 10;
        var okButton = buttonGroup.add("button", undefined, "OK", { name: "ok" });
        var cancelButton = buttonGroup.add("button", undefined, "Cancel", { name: "cancel" });
        
        qtyField.active = true;
        
        // Show dialog and get result
        if (dialog.show() === 1) {
            var currentNum, currentDenom, targetNum, targetDenom, qty;
            
            // Parse current scale
            if (currentDropdown.selection.text === "Custom") {
                currentNum = parseFloat(currentRow.num.text);
                currentDenom = parseFloat(currentRow.denom.text);
            } else {
                var currentParts = currentDropdown.selection.text.split(":");
                currentNum = parseFloat(currentParts[0]);
                currentDenom = parseFloat(currentParts[1]);
            }
            
            // Parse target scale
            if (targetDropdown.selection.text === "Custom") {
                targetNum = parseFloat(targetRow.num.text);
                targetDenom = parseFloat(targetRow.denom.text);
            } else {
                var targetParts = targetDropdown.selection.text.split(":");
                targetNum = parseFloat(targetParts[0]);
                targetDenom = parseFloat(targetParts[1]);
            }
            
            // Parse quantity: typed value wins if anything was entered
            var qtyTyped = qtyField.text.replace(/^\s+|\s+$/g, "");
            if (qtyTyped !== "") {
                qty = parseInt(qtyTyped, 10);
            } else {
                qty = parseInt(qtyDropdown.selection.text, 10);
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
