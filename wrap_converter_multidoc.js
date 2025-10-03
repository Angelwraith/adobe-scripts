/*
@METADATA
{
  "name": "Wrap Template Converter (Multi-Doc)",
  "description": "Prepare Wrap Design Files From Templates - Multiple Documents",
  "version": "2.0",
  "target": "illustrator",
  "tags": ["wrap", "template", "processors", "multi-document"]
}
@END_METADATA
*/

#target illustrator

// ============================================================================
// MAIN ENTRY POINT
// ============================================================================
try {
    if (app.documents.length == 0) {
        alert("Please open at least one document first.");
    } else {
        showDocumentSelectionDialog();
    }
} catch (e) {
    alert("Startup error: " + e.toString());
}

// ============================================================================
// DOCUMENT SELECTION DIALOG
// ============================================================================
function showDocumentSelectionDialog() {
    var dialog = new Window("dialog", "Select Documents to Process");
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.spacing = 15;
    dialog.margins = 20;
    dialog.preferredSize.width = 500;
    
    var titleText = dialog.add("statictext", undefined, "Select wrap template documents to process:");
    titleText.graphics.font = ScriptUI.newFont("dialog", "Bold", 14);
    
    dialog.add("panel");
    
    // Save location section
    var saveLocationGroup = dialog.add("group");
    saveLocationGroup.orientation = "column";
    saveLocationGroup.alignChildren = "fill";
    saveLocationGroup.spacing = 10;
    
    var saveLocationLabel = saveLocationGroup.add("statictext", undefined, "PDF Save Location:");
    saveLocationLabel.graphics.font = ScriptUI.newFont("dialog", "Bold", 12);
    
    var pathInputGroup = saveLocationGroup.add("group");
    pathInputGroup.orientation = "row";
    pathInputGroup.alignChildren = "center";
    pathInputGroup.spacing = 10;
    
    var pathInput = pathInputGroup.add("edittext", undefined, "");
    pathInput.preferredSize.width = 350;
    pathInput.helpTip = "Paste or type the full folder path here, or use Browse button";
    
    var browseButton = pathInputGroup.add("button", undefined, "Browse...");
    browseButton.preferredSize.width = 80;
    
    var selectedFolder = null;
    
    browseButton.onClick = function() {
        var folder = Folder.selectDialog("Select folder to save PDFs:");
        if (folder) {
            selectedFolder = folder;
            pathInput.text = folder.fsName;
        }
    };
    
    // Update selectedFolder when user types or pastes in the text field
    pathInput.onChanging = function() {
        if (pathInput.text && pathInput.text.length > 0) {
            selectedFolder = new Folder(pathInput.text);
        } else {
            selectedFolder = null;
        }
    };
    
    // Prefix section
    var prefixGroup = saveLocationGroup.add("group");
    prefixGroup.orientation = "row";
    prefixGroup.alignChildren = "center";
    prefixGroup.spacing = 10;
    
    var prefixLabel = prefixGroup.add("statictext", undefined, "Filename Prefix:");
    prefixLabel.preferredSize.width = 100;
    
    var prefixInput = prefixGroup.add("edittext", undefined, "");
    prefixInput.preferredSize.width = 330;
    prefixInput.helpTip = "Optional: Add a prefix to all PDF filenames (e.g., 'PROOF')";
    
    dialog.add("panel");
    
    // Document list with checkboxes
    var documentsGroup = dialog.add("group");
    documentsGroup.orientation = "column";
    documentsGroup.alignChildren = "fill";
    
    var docsLabel = documentsGroup.add("statictext", undefined, "Select Documents:");
    docsLabel.graphics.font = ScriptUI.newFont("dialog", "Bold", 12);
    
    var listPanel = documentsGroup.add("panel");
    listPanel.orientation = "column";
    listPanel.alignChildren = "fill";
    listPanel.margins = 10;
    listPanel.preferredSize.height = 180;
    
    var checkboxes = [];
    for (var i = 0; i < app.documents.length; i++) {
        var doc = app.documents[i];
        var checkboxGroup = listPanel.add("group");
        checkboxGroup.orientation = "row";
        checkboxGroup.alignChildren = "left";
        
        var checkbox = checkboxGroup.add("checkbox", undefined, doc.name);
        checkbox.value = true;
        checkbox.preferredSize.width = 400;
        
        checkboxes.push({
            checkbox: checkbox,
            document: doc
        });
    }
    
    dialog.add("panel");
    
    // Selection buttons
    var selectionGroup = dialog.add("group");
    selectionGroup.alignment = "center";
    selectionGroup.spacing = 10;
    
    var selectAllButton = selectionGroup.add("button", undefined, "Select All");
    selectAllButton.preferredSize.width = 80;
    var selectCurrentButton = selectionGroup.add("button", undefined, "Select Current");
    selectCurrentButton.preferredSize.width = 90;
    var selectNoneButton = selectionGroup.add("button", undefined, "Select None");
    selectNoneButton.preferredSize.width = 80;
    
    selectAllButton.onClick = function() {
        for (var i = 0; i < checkboxes.length; i++) {
            checkboxes[i].checkbox.value = true;
        }
    };
    
    selectCurrentButton.onClick = function() {
        var activeDocName = app.activeDocument.name;
        for (var i = 0; i < checkboxes.length; i++) {
            checkboxes[i].checkbox.value = (checkboxes[i].document.name === activeDocName);
        }
    };
    
    selectNoneButton.onClick = function() {
        for (var i = 0; i < checkboxes.length; i++) {
            checkboxes[i].checkbox.value = false;
        }
    };
    
    dialog.add("panel");
    
    // Action buttons
    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "center";
    buttonGroup.spacing = 10;
    
    var cancelButton = buttonGroup.add("button", undefined, "Cancel");
    var processButton = buttonGroup.add("button", undefined, "Process Documents");
    
    var result = null;
    
    cancelButton.onClick = function() {
        result = null;
        dialog.close();
    };
    
    processButton.onClick = function() {
        var selectedDocuments = [];
        for (var i = 0; i < checkboxes.length; i++) {
            if (checkboxes[i].checkbox.value) {
                selectedDocuments.push(checkboxes[i].document);
            }
        }
        
        if (selectedDocuments.length === 0) {
            alert("Please select at least one document to process.");
            return;
        }
        
        // Validate folder path if provided
        if (pathInput.text && pathInput.text.length > 0) {
            selectedFolder = new Folder(pathInput.text);
            if (!selectedFolder.exists) {
                alert("The specified folder does not exist:\n" + pathInput.text + "\n\nPlease check the path and try again.");
                return;
            }
        }
        
        result = {
            documents: selectedDocuments,
            saveFolder: selectedFolder,
            prefix: prefixInput.text
        };
        dialog.close();
    };
    
    dialog.show();
    
    if (result && result.documents && result.documents.length > 0) {
        processDocuments(result.documents, result.saveFolder, result.prefix);
    }
}

// ============================================================================
// DOCUMENT PROCESSING LOOP
// ============================================================================
function processDocuments(documents, saveFolder, prefix) {
    try {
        var overallStartTime = new Date().getTime();
        var successCount = 0;
        var errorCount = 0;
        var errorMessages = [];
        
        // Store document names instead of references
        var docNames = [];
        for (var i = 0; i < documents.length; i++) {
            docNames.push(documents[i].name);
        }
        
        // Process each document
        for (var docIndex = 0; docIndex < docNames.length; docIndex++) {
            var docName = docNames[docIndex];
            
            var docStartTime = new Date().getTime();
            
            // Find and activate document by name
            var targetDoc = null;
            
            for (var activateIndex = 0; activateIndex < app.documents.length; activateIndex++) {
                if (app.documents[activateIndex].name === docName) {
                    targetDoc = app.documents[activateIndex];
                    app.activeDocument = targetDoc;
                    break;
                }
            }
            
            if (!targetDoc) {
                errorMessages.push("Could not find document: " + docName);
                errorCount++;
                continue;
            }
            
            // Force redraw to ensure document is active
            app.redraw();
            $.sleep(100);
            
            // Process the document
            var processResult = processWrapTemplate(targetDoc, saveFolder, prefix);
            
            if (processResult.success) {
                successCount++;
            } else {
                errorCount++;
                errorMessages.push(docName + ": " + processResult.error);
            }
            
            // Pause between documents
            $.sleep(500);
            $.gc();
            app.redraw();
        }
        
        var totalTime = new Date().getTime() - overallStartTime;
        
        // Show final results only
        var resultMessage = "Processing complete!\n\n";
        resultMessage += "Successfully processed: " + successCount + " document(s)\n";
        if (errorCount > 0) {
            resultMessage += "Failed/Skipped: " + errorCount + " document(s)\n\n";
            if (errorMessages.length > 0) {
                resultMessage += "Issues:\n" + errorMessages.join("\n");
            }
        }
        resultMessage += "\n\nTotal time: " + Math.round(totalTime / 1000) + " seconds";
        
        alert(resultMessage);
        
    } catch (error) {
        alert("Error in processing: " + error.toString());
    }
}

// ============================================================================
// WRAP TEMPLATE PROCESSING FUNCTION
// ============================================================================
function processWrapTemplate(doc, saveFolder, prefix) {
    var result = {
        success: false,
        error: ""
    };
    
    // Verify document exists at the start
    if (!doc || !doc.name) {
        result.error = "Invalid document reference";
        return result;
    }
    
    try {
        // Store document name for reference
        var docName = doc.name;
        
        // Check if document is already CMYK - if not, warn and skip
        if (doc.documentColorSpace !== DocumentColorSpace.CMYK) {
            result.error = "Document is not in CMYK color mode. Please convert to CMYK manually first (File > Document Color Mode > CMYK Color)";
            return result;
        }
        
        // 1. Delete "Your Logo Here Image" - it's always the top layer
        if (doc.layers.length > 0) {
            var topLayer = doc.layers[0];
            if (topLayer.name.indexOf("Your Logo Here") !== -1 || topLayer.name.indexOf("Logo") !== -1) {
                topLayer.remove();
            }
        }
        
        // 2. Find and process "Design on Layers Below" layer
        var designLayer = null;
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name.indexOf("Design on Layers Below") !== -1) {
                designLayer = doc.layers[i];
                break;
            }
        }
        
        if (designLayer) {
            // Clear all contents of this layer (including "Design Layer Image")
            clearLayerContents(designLayer);
            
            // Rename the layer to "Design"
            designLayer.name = "Design";
        }
        
        // 3. Rename artboard to "Proof"
        if (doc.artboards.length > 0) {
            doc.artboards[0].name = "Proof";
        }
        
        // 4. Lock "Template Layers | Do not edit!"
        var templateLayer = findLayerByName(doc.layers, "Template Layers | Do not edit!");
        if (templateLayer) {
            templateLayer.locked = true;
        }
        
        // 5. Set active layer to "Design"
        var newDesignLayer = findLayerByName(doc.layers, "Design");
        if (newDesignLayer) {
            doc.activeLayer = newDesignLayer;
        }
        
        // 6. Save as PDF with layers preserved
        var saveResult = saveAsPDFWithLayers(doc, saveFolder, prefix);
        if (!saveResult.success) {
            result.error = saveResult.error;
            return result;
        }
        
        // 7. Documents left open - user can close manually
        
        result.success = true;
        
    } catch (error) {
        result.error = error.message;
    }
    
    return result;
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

// Clear all contents of a layer (sublayers and page items)
function clearLayerContents(layer) {
    try {
        // Remove all sublayers
        if (layer.layers && layer.layers.length > 0) {
            for (var i = layer.layers.length - 1; i >= 0; i--) {
                layer.layers[i].remove();
            }
        }
        
        // Remove all page items (shapes, images, text, etc.)
        if (layer.pageItems && layer.pageItems.length > 0) {
            for (var j = layer.pageItems.length - 1; j >= 0; j--) {
                layer.pageItems[j].remove();
            }
        }
        
        return true;
    } catch (e) {
        return false;
    }
}

// Find layer by name (searches recursively)
function findLayerByName(layersCollection, targetName) {
    for (var i = 0; i < layersCollection.length; i++) {
        var layer = layersCollection[i];
        
        if (layer.name === targetName) {
            return layer;
        }
        
        if (layer.layers && layer.layers.length > 0) {
            var found = findLayerByName(layer.layers, targetName);
            if (found) {
                return found;
            }
        }
    }
    return null;
}

// Save document as PDF using Illustrator Default preset
function saveAsPDFWithLayers(doc, saveFolder, prefix) {
    var result = {
        success: false,
        error: ""
    };
    
    try {
        var pdfFile;
        var docName = doc.name;
        
        // Remove file extension from document name
        var baseName = docName.replace(/\.[^\/\.]+$/, "");
        
        // Extract just the last part after the final underscore for the filename
        // Example: "2015 Ford Transit HR Van 148 WB Extended_Driver" becomes "Driver"
        var shortName = baseName;
        var lastUnderscore = baseName.lastIndexOf("_");
        if (lastUnderscore !== -1) {
            shortName = baseName.substring(lastUnderscore + 1);
        }
        
        // Add prefix if provided
        var finalName = shortName;
        if (prefix && prefix.length > 0) {
            finalName = prefix + "_" + shortName;
        }
        
        // Determine save location
        if (saveFolder) {
            // Use the specified folder
            pdfFile = new File(saveFolder.fsName + "/" + finalName + ".pdf");
        } else {
            // Try to use document's current location
            try {
                var docPath = doc.path;
                pdfFile = new File(docPath + "/" + finalName + ".pdf");
            } catch (e) {
                // Document hasn't been saved yet, prompt user for location
                pdfFile = File.saveDialog("Save PDF as:", "*.pdf");
                
                if (!pdfFile) {
                    result.error = "PDF save cancelled by user";
                    return result;
                }
                
                // Ensure .pdf extension
                if (pdfFile.name.indexOf(".pdf") === -1) {
                    pdfFile = new File(pdfFile.fsName + ".pdf");
                }
            }
        }
        
        // Use the Illustrator Default preset which preserves layers
        var pdfSaveOptions = new PDFSaveOptions();
        pdfSaveOptions.pdfPreset = "[Illustrator Default]";
        
        // Save the PDF
        doc.saveAs(pdfFile, pdfSaveOptions);
        
        result.success = true;
        
    } catch (error) {
        result.error = "Error saving PDF: " + error.message;
    }
    
    return result;
}