#target illustrator

/*@METADATA{
  "name": "Wrap Processor",
  "description": "Export proofs/prints or convert wrap templates with proof-based path prefilling",
  "version": "1.0",
  "target": "illustrator",
  "tags": ["export", "artboards", "jpg", "proof", "print", "wrap", "template"]
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
// FIND PROOF DOCUMENTS
// ============================================================================
function findProofDocuments() {
    var proofDocs = [];
    for (var i = 0; i < app.documents.length; i++) {
        var doc = app.documents[i];
        var docName = doc.name;
        
        // Check if document name contains "_Proof"
        if (docName.match(/_Proof/i)) {
            proofDocs.push({
                document: doc,
                name: docName
            });
        }
    }
    return proofDocs;
}

function findProofFile(folderPath) {
    var folder = Folder(folderPath);
    var files = folder.getFiles("*.pdf");
    
    for (var i = 0; i < files.length; i++) {
        if (files[i] instanceof File) {
            var fileName = files[i].name;
            var decodedFileName = decodeURI(fileName);
            
            if (decodedFileName.match(/\s*_Proof\.pdf$/i)) {
                return files[i];
            }
        }
    }
    
    return null;
}

// ============================================================================
// MODE SELECTION DIALOG
// ============================================================================
function showModeSelectionDialog() {
    // Find proof documents first
    var proofDocuments = findProofDocuments();
    
    var dialog = new Window("dialog", "Export & Convert");
    dialog.orientation = "column";
    dialog.alignChildren = "center";
    dialog.spacing = 20;
    dialog.margins = 30;
    
    var titleText = dialog.add("statictext", undefined, "Select operation:");
    titleText.graphics.font = ScriptUI.newFont("dialog", "Bold", 14);
    
    // Proof document selection (if any proof documents are open)
    var proofList = null;
    if (proofDocuments.length > 0) {
        dialog.add("panel");
        
        var proofLabel = dialog.add("statictext", undefined, "Base paths on open Proof document:");
        proofLabel.graphics.font = ScriptUI.newFont("dialog", "Bold", 11);
        
        proofList = dialog.add("dropdownlist", undefined, []);
        proofList.preferredSize.width = 350;
        
        proofList.add("item", "None - enter paths manually");
        for (var i = 0; i < proofDocuments.length; i++) {
            proofList.add("item", proofDocuments[i].name);
        }
        proofList.selection = 1; // Default to first proof document if available
    }
    
    dialog.add("panel");
    
    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "column";
    buttonGroup.spacing = 15;
    buttonGroup.alignChildren = "fill";
    
    var proofButton = buttonGroup.add("button", undefined, "Export Proof (72 PPI)");
    proofButton.preferredSize.width = 250;
    proofButton.preferredSize.height = 40;
    
    var printButton = buttonGroup.add("button", undefined, "Export Prints (720 PPI)");
    printButton.preferredSize.width = 250;
    printButton.preferredSize.height = 40;
    
    var convertButton = buttonGroup.add("button", undefined, "Convert Wrap Templates");
    convertButton.preferredSize.width = 250;
    convertButton.preferredSize.height = 40;
    
    var cancelButton = buttonGroup.add("button", undefined, "Cancel");
    cancelButton.preferredSize.width = 250;
    
    var result = null;
    
    proofButton.onClick = function() {
        var selectedProofDoc = null;
        if (proofList && proofList.selection && proofList.selection.index > 0) {
            selectedProofDoc = proofDocuments[proofList.selection.index - 1];
        }
        result = {
            mode: "proof",
            proofDoc: selectedProofDoc
        };
        dialog.close();
    };
    
    printButton.onClick = function() {
        var selectedProofDoc = null;
        if (proofList && proofList.selection && proofList.selection.index > 0) {
            selectedProofDoc = proofDocuments[proofList.selection.index - 1];
        }
        result = {
            mode: "print",
            proofDoc: selectedProofDoc
        };
        dialog.close();
    };
    
    convertButton.onClick = function() {
        var selectedProofDoc = null;
        if (proofList && proofList.selection && proofList.selection.index > 0) {
            selectedProofDoc = proofDocuments[proofList.selection.index - 1];
        }
        result = {
            mode: "convert",
            proofDoc: selectedProofDoc
        };
        dialog.close();
    };
    
    cancelButton.onClick = function() {
        result = null;
        dialog.close();
    };
    
    dialog.show();
    
    if (result) {
        if (result.mode === "proof") {
            showExportDialog("proof", 72, result.proofDoc);
        } else if (result.mode === "print") {
            showExportDialog("print", 720, result.proofDoc);
        } else if (result.mode === "convert") {
            showConvertDialog(result.proofDoc);
        }
    }
}

// ============================================================================
// EXPORT DIALOG (FOR PROOF AND PRINT MODES)
// ============================================================================
function showExportDialog(mode, resolution, selectedProofDoc) {
    var folderName = (mode === "proof") ? "Proof" : "PRINT";
    var dialogTitle = (mode === "proof") ? "Export Proof - Select Documents" : "Export Prints - Select Documents";
    
    // Extract path and prefix from selected proof document
    var prefillData = null;
    
    if (selectedProofDoc) {
        var proofDoc = selectedProofDoc.document;
        try {
            if (proofDoc.fullName) {
                var proofPath = proofDoc.fullName;
                var proofFolder = proofPath.parent;
                
                // Find the proof file in the folder (same method as PRIME script)
                var proofFile = findProofFile(proofFolder.fsName);
                var prefix = "";
                
                if (proofFile) {
                    // Extract prefix from proof filename
                    var originalName = decodeURI(proofFile.name);
                    var nameWithoutExtension = originalName.replace(/\.[^\.]+$/, '');
                    prefix = nameWithoutExtension.replace(/\s*_Proof$/i, "");
                }
                
                prefillData = {
                    folder: proofFolder.fsName,
                    prefix: prefix
                };
            }
        } catch (e) {
            // Couldn't get path from proof, leave prefillData null
        }
    }
    
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
        
        // Uncheck if this is the selected proof document
        if (selectedProofDoc && doc.name === selectedProofDoc.name) {
            checkbox.value = false;
        } else {
            checkbox.value = true;
        }
        
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
    
    // Export path and prefix options
    var pathGroup = dialog.add("group");
    pathGroup.orientation = "column";
    pathGroup.alignChildren = "fill";
    pathGroup.spacing = 10;
    
    var pathLabel = pathGroup.add("statictext", undefined, "Export folder path:");
    pathLabel.graphics.font = ScriptUI.newFont("dialog", "Bold", 11);
    
    var exportPathField = pathGroup.add("edittext", undefined, prefillData ? prefillData.folder : "");
    exportPathField.preferredSize.width = 450;
    exportPathField.helpTip = "Path to the folder where exports will be saved";
    
    var prefixLabel = pathGroup.add("statictext", undefined, "Filename prefix:");
    prefixLabel.graphics.font = ScriptUI.newFont("dialog", "Bold", 11);
    
    var prefixField = pathGroup.add("edittext", undefined, prefillData ? prefillData.prefix : "");
    prefixField.preferredSize.width = 450;
    prefixField.helpTip = "Prefix for exported filenames (e.g., 'JobName' creates 'JobName_Side_ArtboardName.jpg')";
    
    dialog.add("panel");
    
    // Info text
    var infoText = dialog.add("statictext", undefined, "Export to: " + folderName + " folder at " + resolution + " PPI");
    infoText.alignment = "center";
    
    dialog.add("panel");
    
    // After export options
    var afterExportGroup = dialog.add("group");
    afterExportGroup.orientation = "column";
    afterExportGroup.alignChildren = "left";
    afterExportGroup.spacing = 5;
    
    var afterExportLabel = afterExportGroup.add("statictext", undefined, "After export:");
    afterExportLabel.graphics.font = ScriptUI.newFont("dialog", "Bold", 11);
    
    var closeRadio = afterExportGroup.add("radiobutton", undefined, "Close without saving");
    closeRadio.value = true;
    var keepOpenRadio = afterExportGroup.add("radiobutton", undefined, "Keep the files open");
    
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
        
        // Get export path and prefix
        var exportPath = exportPathField.text.replace(/^\s+|\s+$/g, '');
        var filePrefix = prefixField.text.replace(/^\s+|\s+$/g, '');
        
        if (exportPath === "") {
            alert("Please enter an export folder path.");
            return;
        }
        
        result = {
            documents: selectedDocuments,
            mode: mode,
            folderName: folderName,
            resolution: resolution,
            closeAfterExport: closeRadio.value,
            exportPath: exportPath,
            filePrefix: filePrefix
        };
        dialog.close();
    };
    
    dialog.show();
    
    if (result && result.documents && result.documents.length > 0) {
        processExportDocuments(result.documents, result.mode, result.folderName, result.resolution, result.closeAfterExport, result.exportPath, result.filePrefix);
    }
}

// ============================================================================
// CONVERT DIALOG (FOR WRAP TEMPLATE CONVERSION)
// ============================================================================
function showConvertDialog(selectedProofDoc) {
    // Extract path and prefix from selected proof document
    var prefillData = null;
    
    if (selectedProofDoc) {
        var proofDoc = selectedProofDoc.document;
        try {
            if (proofDoc.fullName) {
                var proofPath = proofDoc.fullName;
                var proofFolder = proofPath.parent;
                
                // Find the proof file in the folder
                var proofFile = findProofFile(proofFolder.fsName);
                var prefix = "";
                
                if (proofFile) {
                    var originalName = decodeURI(proofFile.name);
                    var nameWithoutExtension = originalName.replace(/\.[^\.]+$/, '');
                    prefix = nameWithoutExtension.replace(/\s*_Proof$/i, "");
                }
                
                prefillData = {
                    folder: proofFolder.fsName,
                    prefix: prefix
                };
            }
        } catch (e) {
            // Couldn't get path from proof, leave prefillData null
        }
    }
    
    var dialog = new Window("dialog", "Convert Wrap Templates");
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.spacing = 15;
    dialog.margins = 20;
    dialog.preferredSize.width = 500;
    
    var titleText = dialog.add("statictext", undefined, "Select wrap template documents to convert:");
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
    
    var pathInput = pathInputGroup.add("edittext", undefined, prefillData ? prefillData.folder : "");
    pathInput.preferredSize.width = 350;
    pathInput.helpTip = "Paste or type the full folder path here, or use Browse button";
    
    var browseButton = pathInputGroup.add("button", undefined, "Browse...");
    browseButton.preferredSize.width = 80;
    
    var selectedFolder = prefillData ? new Folder(prefillData.folder) : null;
    
    browseButton.onClick = function() {
        var folder = Folder.selectDialog("Select folder to save PDFs:");
        if (folder) {
            selectedFolder = folder;
            pathInput.text = folder.fsName;
        }
    };
    
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
    
    var prefixInput = prefixGroup.add("edittext", undefined, prefillData ? prefillData.prefix : "");
    prefixInput.preferredSize.width = 330;
    prefixInput.helpTip = "Optional: Add a prefix to all PDF filenames";
    
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
        
        // Uncheck if this is the selected proof document
        if (selectedProofDoc && doc.name === selectedProofDoc.name) {
            checkbox.value = false;
        } else {
            checkbox.value = true;
        }
        
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
        processConvertDocuments(result.documents, result.saveFolder, result.prefix);
    }
}

// ============================================================================
// EXPORT DOCUMENT PROCESSING LOOP
// ============================================================================
function processExportDocuments(documents, mode, folderName, resolution, closeAfterExport, customExportPath, filePrefix) {
    try {
        var overallStartTime = new Date().getTime();
        var totalExported = 0;
        var exportSummary = [];
        
        // Store document file paths instead of references
        var docPaths = [];
        for (var i = 0; i < documents.length; i++) {
            // Only check if fullName exists (not saved status, as PDFs may show as unsaved)
            if (documents[i].fullName && documents[i].fullName.toString() !== "") {
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
            var exportResult = processExportDocument(targetDoc, mode, folderName, resolution, customExportPath, filePrefix);
            if (exportResult && exportResult.count > 0) {
                totalExported += exportResult.count;
                exportSummary.push(docInfo.name + ": " + exportResult.count + " artboards exported");
            }
            
            // Handle post-export actions
            if (closeAfterExport) {
                try {
                    targetDoc.close(SaveOptions.DONOTSAVECHANGES);
                } catch (closeError) {
                    alert("Warning: Could not close document '" + docInfo.name + "'.\n" + closeError.toString());
                }
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
// EXPORT SINGLE DOCUMENT
// ============================================================================
function processExportDocument(doc, mode, folderName, resolution, customExportPath, filePrefix) {
    try {
        // Double-check document has a file path
        if (!doc.fullName || doc.fullName.toString() === "") {
            alert("Document '" + doc.name + "' has no file path. Please save it first. Skipping.");
            return { count: 0 };
        }
        
        // Get document path and name
        var docPath = doc.fullName;
        var docFolder = docPath.parent;
        var docName = doc.name.replace(/\.[^\.]+$/, '');
        var docFullPath = doc.fullName.fsName;
        
        // Navigate to export folder
        var exportFolder;
        if (customExportPath && customExportPath !== "") {
            exportFolder = new Folder(customExportPath + "/" + folderName);
        } else {
            if (mode === "proof") {
                exportFolder = docFolder.parent;
                exportFolder = new Folder(exportFolder.fsName + "/" + folderName);
            } else {
                exportFolder = docFolder.parent.parent;
                exportFolder = new Folder(exportFolder.fsName + "/" + folderName);
            }
        }
        
        if (!exportFolder.exists) {
            exportFolder.create();
        }
        
        app.preferences.setIntegerPreference('plugin/SmartExportUI/CreateFoldersPreference', 0);
        
        // Collect artboards to export
        var artboardsToExport = [];
        for (var i = 0; i < doc.artboards.length; i++) {
            var artboard = doc.artboards[i];
            var artboardName = artboard.name;
            
            if (mode === "proof") {
                if (artboardName.toLowerCase().indexOf("proof") !== -1) {
                    artboardsToExport.push({
                        index: i,
                        name: artboardName
                    });
                }
            } else if (mode === "print") {
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
            
            // Re-find document before each export
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
            var exportPrefix = docName;
            
            if (filePrefix && filePrefix !== "") {
                var docNameNoExt = docName;
                
                if (docNameNoExt.indexOf(filePrefix) === 0) {
                    var suffix = docNameNoExt.substring(filePrefix.length);
                    exportPrefix = filePrefix + suffix;
                } else {
                    exportPrefix = docName;
                }
            }
            
            exportArtboard(currentDoc, artboardInfo.index, artboardInfo.name, exportPrefix, exportFolder, resolution);
            exportCount++;
            
            $.sleep(100);
        }
        
        return { count: exportCount };
        
    } catch (error) {
        alert("Error processing document '" + doc.name + "':\n" + error.toString());
        return { count: 0 };
    }
}

// ============================================================================
// EXPORT SINGLE ARTBOARD
// ============================================================================
function exportArtboard(doc, artboardIndex, artboardName, docName, destFolder, resolution) {
    try {
        var exportItem = new ExportForScreensItemToExport();
        exportItem.artboards = String(artboardIndex + 1);
        exportItem.document = false;
        
        var jpegOptions = new ExportForScreensOptionsJPEG();
        jpegOptions.compressionMethod = JPEGCompressionMethodType.BASELINESTANDARD;
        jpegOptions.antiAliasing = AntiAliasingMethod.TYPEOPTIMIZED;
        jpegOptions.qualitySetting = 100;
        jpegOptions.scaleType = ExportForScreensScaleType.SCALEBYRESOLUTION;
        jpegOptions.scaleTypeValue = resolution;
        
        var originalName = doc.artboards[artboardIndex].name;
        var desiredFileName = docName + "_" + artboardName;
        doc.artboards[artboardIndex].name = desiredFileName;
        
        doc.exportForScreens(destFolder, ExportForScreensType.SE_JPEG100, jpegOptions, exportItem);
        
        doc.artboards[artboardIndex].name = originalName;
        
    } catch (error) {
        try {
            doc.artboards[artboardIndex].name = originalName;
        } catch (e) {}
        throw error;
    }
}

// ============================================================================
// CONVERT DOCUMENT PROCESSING LOOP
// ============================================================================
function processConvertDocuments(documents, saveFolder, prefix) {
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
            
            // Find document by name
            var targetDoc = null;
            for (var searchIndex = 0; searchIndex < app.documents.length; searchIndex++) {
                if (app.documents[searchIndex].name === docName) {
                    targetDoc = app.documents[searchIndex];
                    app.activeDocument = targetDoc;
                    break;
                }
            }
            
            if (!targetDoc) {
                errorMessages.push(docName + ": Document not found");
                errorCount++;
                continue;
            }
            
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
            
            $.sleep(500);
            $.gc();
            app.redraw();
        }
        
        var totalTime = new Date().getTime() - overallStartTime;
        
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
    
    if (!doc || !doc.name) {
        result.error = "Invalid document reference";
        return result;
    }
    
    try {
        var docName = doc.name;
        
        if (doc.documentColorSpace !== DocumentColorSpace.CMYK) {
            result.error = "Document is not in CMYK color mode. Please convert to CMYK manually first (File > Document Color Mode > CMYK Color)";
            return result;
        }
        
        // 1. Delete "Your Logo Here Image" - top layer
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
            clearLayerContents(designLayer);
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
        
        // 6. Save as PDF
        var saveResult = saveAsPDFWithLayers(doc, saveFolder, prefix);
        if (!saveResult.success) {
            result.error = saveResult.error;
            return result;
        }
        
        result.success = true;
        
    } catch (error) {
        result.error = error.message;
    }
    
    return result;
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

function clearLayerContents(layer) {
    try {
        if (layer.layers && layer.layers.length > 0) {
            for (var i = layer.layers.length - 1; i >= 0; i--) {
                layer.layers[i].remove();
            }
        }
        
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

function saveAsPDFWithLayers(doc, saveFolder, prefix) {
    var result = {
        success: false,
        error: ""
    };
    
    try {
        var pdfFile;
        var docName = doc.name;
        
        var baseName = docName.replace(/\.[^\/\.]+$/, "");
        
        // Extract just the last part after the final underscore
        var shortName = baseName;
        var lastUnderscore = baseName.lastIndexOf("_");
        if (lastUnderscore !== -1) {
            shortName = baseName.substring(lastUnderscore + 1);
        }
        
        var finalName = shortName;
        if (prefix && prefix.length > 0) {
            finalName = prefix + "_" + shortName;
        }
        
        if (saveFolder) {
            pdfFile = new File(saveFolder.fsName + "/" + finalName + ".pdf");
        } else {
            try {
                var docPath = doc.path;
                pdfFile = new File(docPath + "/" + finalName + ".pdf");
            } catch (e) {
                pdfFile = File.saveDialog("Save PDF as:", "*.pdf");
                
                if (!pdfFile) {
                    result.error = "PDF save cancelled by user";
                    return result;
                }
                
                if (pdfFile.name.indexOf(".pdf") === -1) {
                    pdfFile = new File(pdfFile.fsName + ".pdf");
                }
            }
        }
        
        var pdfSaveOptions = new PDFSaveOptions();
        pdfSaveOptions.pdfPreset = "[Illustrator Default]";
        
        doc.saveAs(pdfFile, pdfSaveOptions);
        
        result.success = true;
        
    } catch (error) {
        result.error = "Error saving PDF: " + error.message;
    }
    
    return result;
}
