#target illustrator

/*@METADATA{
  "name": "Export Artboards - Proof or Print",
  "description": "Exports artboards as proofs (72 PPI) or prints (720 PPI) to respective folders",
  "version": "5.0",
  "target": "illustrator",
  "tags": ["export", "artboards", "jpg", "proof", "print"]
}@END_METADATA*/

// ============================================================================
// MAIN ENTRY POINT
// ============================================================================
try {
    if (app.documents.length == 0) {
        alert("Please open at least one document first.");
    } else {
        showModeSelectionDialog();
    }
} catch (e) {
    alert("Startup error: " + e.toString());
}

// ============================================================================
// MODE SELECTION DIALOG (PROOF OR PRINT)
// ============================================================================
function showModeSelectionDialog() {
    var dialog = new Window("dialog", "Export Mode");
    dialog.orientation = "column";
    dialog.alignChildren = "center";
    dialog.spacing = 20;
    dialog.margins = 30;
    
    var titleText = dialog.add("statictext", undefined, "Select export mode:");
    titleText.graphics.font = ScriptUI.newFont("dialog", "Bold", 14);
    
    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "column";
    buttonGroup.spacing = 15;
    buttonGroup.alignChildren = "fill";
    
    var proofButton = buttonGroup.add("button", undefined, "Export Proof");
    proofButton.preferredSize.width = 200;
    proofButton.preferredSize.height = 40;
    
    var printButton = buttonGroup.add("button", undefined, "Export Prints");
    printButton.preferredSize.width = 200;
    printButton.preferredSize.height = 40;
    
    var cancelButton = buttonGroup.add("button", undefined, "Cancel");
    cancelButton.preferredSize.width = 200;
    
    var result = null;
    
    proofButton.onClick = function() {
        result = "proof";
        dialog.close();
    };
    
    printButton.onClick = function() {
        result = "print";
        dialog.close();
    };
    
    cancelButton.onClick = function() {
        result = null;
        dialog.close();
    };
    
    dialog.show();
    
    if (result === "proof") {
        showDocumentSelectionDialog("proof", 72);
    } else if (result === "print") {
        showDocumentSelectionDialog("print", 720);
    }
}

// ============================================================================
// DOCUMENT SELECTION DIALOG
// ============================================================================
function showDocumentSelectionDialog(mode, resolution) {
    var folderName = (mode === "proof") ? "Proof" : "PRINT";
    var dialogTitle = (mode === "proof") ? "Export Proof - Select Documents" : "Export Prints - Select Documents";
    
    var dialog = new Window("dialog", dialogTitle);
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.spacing = 15;
    dialog.margins = 20;
    dialog.preferredSize.width = 500;
    
    var titleText = dialog.add("statictext", undefined, "Select documents to export:");
    titleText.graphics.font = ScriptUI.newFont("dialog", "Bold", 14);
    
    dialog.add("panel");
    
    // Document list with checkboxes
    var documentsGroup = dialog.add("group");
    documentsGroup.orientation = "column";
    documentsGroup.alignChildren = "fill";
    
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
    
    // Info text
    var infoText = dialog.add("statictext", undefined, "Export to: " + folderName + " folder at " + resolution + " PPI");
    infoText.alignment = "center";
    
    dialog.add("panel");
    
    // Action buttons
    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "center";
    buttonGroup.spacing = 10;
    
    var cancelButton = buttonGroup.add("button", undefined, "Cancel");
    var processButton = buttonGroup.add("button", undefined, "Export Documents");
    
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
        
        result = {
            documents: selectedDocuments,
            mode: mode,
            folderName: folderName,
            resolution: resolution
        };
        dialog.close();
    };
    
    dialog.show();
    
    if (result && result.documents && result.documents.length > 0) {
        processDocuments(result.documents, result.mode, result.folderName, result.resolution);
    }
}

// ============================================================================
// DOCUMENT PROCESSING LOOP
// ============================================================================
function processDocuments(documents, mode, folderName, resolution) {
    try {
        var overallStartTime = new Date().getTime();
        var totalExported = 0;
        var exportSummary = [];
        
        // Store document file paths instead of references (since references can become stale)
        var docPaths = [];
        for (var i = 0; i < documents.length; i++) {
            if (documents[i].saved && documents[i].fullName.toString() !== "") {
                docPaths.push({
                    path: documents[i].fullName.fsName,
                    name: documents[i].name
                });
            }
        }
        
        // Process each document by path
        for (var docIndex = 0; docIndex < docPaths.length; docIndex++) {
            var docInfo = docPaths[docIndex];
            
            var docStartTime = new Date().getTime();
            
            // Find and activate document by full path
            var targetDoc = null;
            for (var searchIndex = 0; searchIndex < app.documents.length; searchIndex++) {
                if (app.documents[searchIndex].fullName.fsName === docInfo.path) {
                    targetDoc = app.documents[searchIndex];
                    app.activeDocument = targetDoc;
                    break;
                }
            }
            
            if (!targetDoc) {
                alert("ERROR: Could not find document: " + docInfo.name);
                continue;
            }
            
            // Force redraw to ensure document is active
            app.redraw();
            $.sleep(100);
            
            // Process this document
            var exportResult = processDocument(targetDoc, mode, folderName, resolution);
            if (exportResult && exportResult.count > 0) {
                totalExported += exportResult.count;
                exportSummary.push(docInfo.name + ": " + exportResult.count + " artboards exported");
            }
            
            // Save the document after export to prevent "unsaved" state issues
            // Use save() instead of saveAs() to keep original format and settings
            try {
                if (!targetDoc.saved) {
                    targetDoc.save();
                }
            } catch (saveError) {
                alert("Warning: Could not save document '" + docInfo.name + "' after export.\n" + saveError.toString());
            }
            
            // Cleanup between documents
            if (docIndex < docPaths.length - 1) {
                try {
                    app.redraw();
                    $.gc();
                    $.sleep(300);
                } catch (cleanupError) {}
            }
        }
        
        var totalTime = new Date().getTime() - overallStartTime;
        
        var summaryMessage = "Export complete!\n\n";
        summaryMessage += "Total artboards exported: " + totalExported + "\n";
        summaryMessage += "Total time: " + Math.round(totalTime / 1000) + " seconds\n\n";
        summaryMessage += exportSummary.join("\n");
        
        alert(summaryMessage);
        
    } catch (error) {
        alert("Error in processing: " + error.toString());
    }
}

// ============================================================================
// DOCUMENT PROCESSING FUNCTION
// ============================================================================
function processDocument(doc, mode, folderName, resolution) {
    try {
        // Double-check document is saved (should be, but verify)
        if (!doc.saved || doc.fullName.toString() === "") {
            alert("Document '" + doc.name + "' is not saved. Skipping.");
            return { count: 0 };
        }
        
        // Get document path and name - store these before any exports
        var docPath = doc.fullName;
        var docFolder = docPath.parent;
        var docName = doc.name.replace(/\.[^\.]+$/, ''); // Remove extension
        var docFullPath = doc.fullName.fsName; // Store full path for re-finding document
        
        // Navigate to export folder
        var exportFolder;
        if (mode === "proof") {
            // Proof folder is inside PRIME folder (1 level up from document location)
            exportFolder = docFolder.parent;
            exportFolder = new Folder(exportFolder.fsName + "/" + folderName);
        } else {
            // PRINT folder is at same level as PRIME (2 levels up from document location)
            exportFolder = docFolder.parent.parent;
            exportFolder = new Folder(exportFolder.fsName + "/" + folderName);
        }
        
        // Create export folder if it doesn't exist
        if (!exportFolder.exists) {
            exportFolder.create();
        }
        
        // Disable subfolder creation for Export for Screens
        app.preferences.setIntegerPreference('plugin/SmartExportUI/CreateFoldersPreference', 0);
        
        // Collect artboards to export BEFORE doing any exports
        var artboardsToExport = [];
        for (var i = 0; i < doc.artboards.length; i++) {
            var artboard = doc.artboards[i];
            var artboardName = artboard.name;
            
            if (mode === "proof") {
                // For proof mode, only export artboards with "proof" in the name (case-insensitive)
                if (artboardName.toLowerCase().indexOf("proof") !== -1) {
                    artboardsToExport.push({
                        index: i,
                        name: artboardName
                    });
                }
            } else if (mode === "print") {
                // For print mode, skip artboards with "proof" in the name (case-insensitive)
                if (artboardName.toLowerCase().indexOf("proof") === -1) {
                    artboardsToExport.push({
                        index: i,
                        name: artboardName
                    });
                }
            }
        }
        
        if (artboardsToExport.length === 0) {
            if (mode === "proof") {
                alert("No proof artboards to export in '" + doc.name + "'.\nSkipping this document.");
            } else {
                alert("No artboards to export in '" + doc.name + "' (all contain 'proof').\nSkipping this document.");
            }
            return { count: 0 };
        }
        
        // Export each artboard
        var exportCount = 0;
        for (var j = 0; j < artboardsToExport.length; j++) {
            var artboardInfo = artboardsToExport[j];
            
            // Re-find document before each export (in case reference becomes stale)
            var currentDoc = null;
            for (var findIndex = 0; findIndex < app.documents.length; findIndex++) {
                if (app.documents[findIndex].fullName.fsName === docFullPath) {
                    currentDoc = app.documents[findIndex];
                    app.activeDocument = currentDoc;
                    break;
                }
            }
            
            if (!currentDoc) {
                alert("ERROR: Lost reference to document during export.");
                break;
            }
            
            // Export this artboard
            exportArtboard(currentDoc, artboardInfo.index, artboardInfo.name, docName, exportFolder, resolution);
            exportCount++;
            
            // Small delay between artboards to let Illustrator catch up
            $.sleep(100);
        }
        
        return { count: exportCount };
        
    } catch (error) {
        alert("Error processing document '" + doc.name + "':\n" + error.toString());
        return { count: 0 };
    }
}

// ============================================================================
// EXPORT SINGLE ARTBOARD WITH RESOLUTION
// ============================================================================
function exportArtboard(doc, artboardIndex, artboardName, docName, destFolder, resolution) {
    try {
        // Create an ExportForScreensItemToExport object
        var exportItem = new ExportForScreensItemToExport();
        exportItem.artboards = String(artboardIndex + 1); // Artboard indices are 1-based
        exportItem.document = false;
        
        // Create JPEG options with specified resolution
        var jpegOptions = new ExportForScreensOptionsJPEG();
        jpegOptions.compressionMethod = JPEGCompressionMethodType.BASELINESTANDARD;
        jpegOptions.antiAliasing = AntiAliasingMethod.TYPEOPTIMIZED;
        jpegOptions.qualitySetting = 100; // Maximum quality
        jpegOptions.scaleType = ExportForScreensScaleType.SCALEBYRESOLUTION;
        jpegOptions.scaleTypeValue = resolution;
        
        // Temporarily rename artboard to get correct filename
        var originalName = doc.artboards[artboardIndex].name;
        var desiredFileName = docName + "_" + artboardName;
        doc.artboards[artboardIndex].name = desiredFileName;
        
        // Perform the export
        doc.exportForScreens(destFolder, ExportForScreensType.SE_JPEG100, jpegOptions, exportItem);
        
        // Restore original artboard name
        doc.artboards[artboardIndex].name = originalName;
        
    } catch (error) {
        // Restore original name even if export fails
        try {
            doc.artboards[artboardIndex].name = originalName;
        } catch (e) {}
        throw error;
    }
}
