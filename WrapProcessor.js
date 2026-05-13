#target illustrator

/*@METADATA{
  "name": "Wrap Processor",
  "description": "Export proofs/prints or convert wrap templates with proof-based path prefilling. Convert mode auto-converts RGB docs to CMYK via BridgeTalk-isolated calls.",
  "version": "2.0",
  "target": "illustrator",
  "tags": ["export", "artboards", "jpg", "proof", "print", "wrap", "template", "cmyk"]
}@END_METADATA*/

// ============================================================================
// MAIN ENTRY POINT
// ============================================================================
try {
    // No doc-required check here — "Open Wrap TIFs" mode needs to run with
    // zero docs open. Modes that need a doc check inside their handlers.
    showModeSelectionDialog();
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
// RECURSIVELY FIND ALL TIF FILES under a folder (including subfolders)
// ============================================================================
function findAllTifs(folder) {
    var tifs = [];
    if (!folder || !folder.exists) return tifs;

    var items = [];
    try { items = folder.getFiles(); } catch (e) { return tifs; }

    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        if (item instanceof Folder) {
            var sub = findAllTifs(item);
            for (var j = 0; j < sub.length; j++) tifs.push(sub[j]);
        } else if (item instanceof File) {
            var name = item.name;
            if (/\.tiff?$/i.test(name)) {
                tifs.push(item);
            }
        }
    }
    return tifs;
}

// ============================================================================
// IS WRAP-SIDE TIF — name contains driver/front/pssngr/passenger/rear/top
// ============================================================================
function isWrapSideTIF(filename) {
    var lower = filename.toLowerCase();
    var keywords = ["driver", "front", "pssngr", "passenger", "rear", "top"];
    for (var i = 0; i < keywords.length; i++) {
        if (lower.indexOf(keywords[i]) !== -1) return true;
    }
    return false;
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

    var openTifsButton = buttonGroup.add("button", undefined, "Open Wrap TIFs");
    openTifsButton.preferredSize.width = 250;
    openTifsButton.preferredSize.height = 40;

    var convertButton = buttonGroup.add("button", undefined, "Convert Wrap Templates");
    convertButton.preferredSize.width = 250;
    convertButton.preferredSize.height = 40;

    var proofButton = buttonGroup.add("button", undefined, "Export Proof (72 PPI)");
    proofButton.preferredSize.width = 250;
    proofButton.preferredSize.height = 40;

    var printButton = buttonGroup.add("button", undefined, "Export Prints (720 PPI)");
    printButton.preferredSize.width = 250;
    printButton.preferredSize.height = 40;

    var cancelButton = buttonGroup.add("button", undefined, "Cancel");
    cancelButton.preferredSize.width = 250;

    var result = null;

    function getSelectedProofDoc() {
        if (proofList && proofList.selection && proofList.selection.index > 0) {
            return proofDocuments[proofList.selection.index - 1];
        }
        return null;
    }

    openTifsButton.onClick = function() {
        result = { mode: "openTifs", proofDoc: getSelectedProofDoc() };
        dialog.close();
    };

    proofButton.onClick = function() {
        if (app.documents.length === 0) {
            alert("Please open at least one document first (or use 'Open Wrap TIFs' to load some).");
            return;
        }
        result = { mode: "proof", proofDoc: getSelectedProofDoc() };
        dialog.close();
    };

    printButton.onClick = function() {
        if (app.documents.length === 0) {
            alert("Please open at least one document first (or use 'Open Wrap TIFs' to load some).");
            return;
        }
        result = { mode: "print", proofDoc: getSelectedProofDoc() };
        dialog.close();
    };

    convertButton.onClick = function() {
        if (app.documents.length === 0) {
            alert("Please open at least one document first (or use 'Open Wrap TIFs' to load some).");
            return;
        }
        result = { mode: "convert", proofDoc: getSelectedProofDoc() };
        dialog.close();
    };

    cancelButton.onClick = function() {
        result = null;
        dialog.close();
    };

    dialog.show();

    if (result) {
        if (result.mode === "openTifs") {
            showOpenTifsDialog(result.proofDoc);
        } else if (result.mode === "proof") {
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
// OPEN WRAP TIFS DIALOG
// Recursively scans the proof file's folder for .tif files, shows them in a
// checkbox list with wrap-side names pre-selected (driver/front/pssngr/rear/
// top). Opens the selected files with userInteractionLevel set to
// DONTDISPLAYALERTS so the TIFF Import Options dialog is auto-dismissed for
// each file instead of requiring an OK click.
// ============================================================================
function showOpenTifsDialog(selectedProofDoc) {
    // Determine starting folder (proof doc's parent if available)
    var startingFolder = null;
    if (selectedProofDoc) {
        try {
            if (selectedProofDoc.document.fullName) {
                startingFolder = selectedProofDoc.document.fullName.parent;
            }
        } catch (e) {}
    }

    var dialog = new Window("dialog", "Open Wrap TIFs");
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.spacing = 15;
    dialog.margins = 20;
    dialog.preferredSize.width = 620;

    var titleText = dialog.add("statictext", undefined, "Select TIF files to open:");
    titleText.graphics.font = ScriptUI.newFont("dialog", "Bold", 14);

    dialog.add("panel");

    // Search location section
    var locGroup = dialog.add("group");
    locGroup.orientation = "column";
    locGroup.alignChildren = "fill";
    locGroup.spacing = 10;

    var locLabel = locGroup.add("statictext", undefined, "Search Folder (includes subfolders):");
    locLabel.graphics.font = ScriptUI.newFont("dialog", "Bold", 12);

    var pathRow = locGroup.add("group");
    pathRow.orientation = "row";
    pathRow.alignChildren = "center";
    pathRow.spacing = 10;

    var pathInput = pathRow.add("edittext", undefined, startingFolder ? startingFolder.fsName : "");
    pathInput.preferredSize.width = 380;

    var browseButton = pathRow.add("button", undefined, "Browse...");
    browseButton.preferredSize.width = 80;

    var refreshButton = pathRow.add("button", undefined, "Refresh");
    refreshButton.preferredSize.width = 70;

    dialog.add("panel");

    // File list section
    var listLabel = dialog.add("statictext", undefined, "TIF Files Found:");
    listLabel.graphics.font = ScriptUI.newFont("dialog", "Bold", 12);

    var listPanel = dialog.add("panel");
    listPanel.orientation = "column";
    listPanel.alignChildren = "left";
    listPanel.margins = 10;
    listPanel.preferredSize.height = 260;

    var checkboxes = [];
    var currentFolder = startingFolder;

    function refreshList() {
        // Remove existing children from listPanel
        while (listPanel.children && listPanel.children.length > 0) {
            try { listPanel.remove(listPanel.children[0]); } catch (e) { break; }
        }
        checkboxes = [];

        if (!currentFolder || !currentFolder.exists) {
            listPanel.add("statictext", undefined, "(Folder not found — pick one with Browse)");
            try { dialog.layout.layout(true); } catch (e) {}
            return;
        }

        var tifs = findAllTifs(currentFolder);
        if (tifs.length === 0) {
            listPanel.add("statictext", undefined, "(No TIF files found in this folder or its subfolders)");
            try { dialog.layout.layout(true); } catch (e) {}
            return;
        }

        // Sort alphabetically by path
        tifs.sort(function(a, b) {
            var an = a.fsName.toLowerCase(), bn = b.fsName.toLowerCase();
            return an < bn ? -1 : (an > bn ? 1 : 0);
        });

        var basePathLen = currentFolder.fsName.length + 1;
        for (var i = 0; i < tifs.length; i++) {
            var tif = tifs[i];
            var label = tif.name;
            // Show relative path if the file is in a subfolder
            try {
                if (tif.fsName.length > basePathLen) {
                    label = tif.fsName.substring(basePathLen);
                }
            } catch (eRel) {}
            // Decode URI-encoded characters (e.g. spaces shown as %20)
            try { label = decodeURI(label); } catch (eDec) {}

            var cb = listPanel.add("checkbox", undefined, label);
            cb._tif = tif;
            cb.value = isWrapSideTIF(tif.name);  // pre-select wrap sides
            cb.preferredSize.width = 560;
            checkboxes.push(cb);
        }

        try { dialog.layout.layout(true); } catch (e) {}
    }

    browseButton.onClick = function() {
        var folder = Folder.selectDialog("Select folder containing TIF files:", currentFolder || undefined);
        if (folder) {
            currentFolder = folder;
            pathInput.text = folder.fsName;
            refreshList();
        }
    };

    refreshButton.onClick = function() {
        var p = pathInput.text;
        if (p && p.length > 0) {
            currentFolder = new Folder(p);
            refreshList();
        }
    };

    // Selection helper buttons
    var selGroup = dialog.add("group");
    selGroup.alignment = "center";
    selGroup.spacing = 10;

    var selAllBtn = selGroup.add("button", undefined, "Select All");
    selAllBtn.preferredSize.width = 80;
    var selNoneBtn = selGroup.add("button", undefined, "Select None");
    selNoneBtn.preferredSize.width = 90;
    var selWrapBtn = selGroup.add("button", undefined, "Select Wrap Sides");
    selWrapBtn.preferredSize.width = 130;

    selAllBtn.onClick = function() {
        for (var i = 0; i < checkboxes.length; i++) checkboxes[i].value = true;
    };
    selNoneBtn.onClick = function() {
        for (var i = 0; i < checkboxes.length; i++) checkboxes[i].value = false;
    };
    selWrapBtn.onClick = function() {
        for (var i = 0; i < checkboxes.length; i++) {
            checkboxes[i].value = isWrapSideTIF(checkboxes[i]._tif.name);
        }
    };

    dialog.add("panel");

    // Action buttons
    var btnRow = dialog.add("group");
    btnRow.alignment = "center";
    btnRow.spacing = 10;
    var cancelBtn = btnRow.add("button", undefined, "Cancel");
    var openOnlyBtn = btnRow.add("button", undefined, "Open Only");
    var openConvertBtn = btnRow.add("button", undefined, "Open & Convert");

    var selectedFiles = null;
    var chosenAction = null; // "open" or "convert"

    cancelBtn.onClick = function() {
        selectedFiles = null;
        dialog.close();
    };

    function collectAndClose(action) {
        var picked = [];
        for (var i = 0; i < checkboxes.length; i++) {
            if (checkboxes[i].value) picked.push(checkboxes[i]._tif);
        }
        if (picked.length === 0) {
            alert("Please check at least one TIF file to open.");
            return;
        }
        selectedFiles = picked;
        chosenAction = action;
        dialog.close();
    }

    openOnlyBtn.onClick    = function() { collectAndClose("open"); };
    openConvertBtn.onClick = function() { collectAndClose("convert"); };

    // Populate the initial list
    refreshList();

    dialog.show();

    if (selectedFiles && selectedFiles.length > 0) {
        openSelectedTifs(selectedFiles);
        if (chosenAction === "convert") {
            // Brief pause to let Illustrator finish processing the opens
            // before the convert dialog enumerates app.documents.
            $.sleep(500);
            showConvertDialog(selectedProofDoc);
        }
    }
}

// ============================================================================
// OPEN SELECTED TIFS (DONTDISPLAYALERTS suppresses the Import Options modal)
// ============================================================================
function openSelectedTifs(tifFiles) {
    var prevLevel = app.userInteractionLevel;
    try { app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS; } catch (eLvl) {}

    var openedCount = 0;
    var errors = [];

    try {
        for (var i = 0; i < tifFiles.length; i++) {
            try {
                app.open(tifFiles[i]);
                openedCount++;
                $.sleep(200);
            } catch (eOpen) {
                errors.push(tifFiles[i].name + ": " + eOpen.message);
            }
        }
    } finally {
        try { app.userInteractionLevel = prevLevel; } catch (eR) {}
    }

    var msg = "Opened " + openedCount + " TIF file" + (openedCount === 1 ? "" : "s") + ".";
    if (errors.length > 0) {
        msg += "\n\nErrors:\n  • " + errors.join("\n  • ");
    }
    alert(msg);
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

    // After-process options
    var afterGroup = dialog.add("group");
    afterGroup.orientation = "column";
    afterGroup.alignChildren = "left";
    afterGroup.spacing = 5;

    var afterLabel = afterGroup.add("statictext", undefined, "After processing:");
    afterLabel.graphics.font = ScriptUI.newFont("dialog", "Bold", 11);
    var closeRadio  = afterGroup.add("radiobutton", undefined, "Close converted documents");
    closeRadio.value = true;
    var reopenRadio = afterGroup.add("radiobutton", undefined, "Reopen the saved PDFs");

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
            prefix: prefixInput.text,
            reopenPdfs: reopenRadio.value
        };
        dialog.close();
    };

    dialog.show();

    if (result && result.documents && result.documents.length > 0) {
        processConvertDocuments(result.documents, result.saveFolder, result.prefix, result.reopenPdfs);
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
// CONVERT DOCUMENT PROCESSING LOOP (v2.0 — BridgeTalk-isolated CMYK conversion)
// ============================================================================
// The CMYK conversion uses app.executeMenuCommand("doc-color-cmyk"), which in
// this Illustrator install throws a spurious "there is no document" error
// that poisons the scripting bridge for the entire script execution context.
//
// Workaround: every Illustrator API call lives inside a BridgeTalk message.
// Each BT body runs in a fresh script context, so the broken bridge from one
// message can't leak into the next. The main script never touches app.*
// during iteration — it only orchestrates BT calls and reads back result
// strings.
//
// Per file:
//   BT1: find doc by name, activate, run doc-color-cmyk if RGB
//   BT2: find doc by name, do the wrap-template work (delete logo layer,
//        clear/rename design layer, rename artboard, lock template, set
//        active layer), saveAs PDF, close
// ============================================================================
function processConvertDocuments(documents, saveFolder, prefix, reopenPdfs) {
    if (!saveFolder) {
        alert("Please specify a save folder before processing.");
        return;
    }

    // Create / verify the Design subfolder inside the user-selected save folder.
    // All converted PDFs go into <saveFolder>/Design/.
    var designFolder = new Folder(saveFolder.fsName + "/Design");
    if (!designFolder.exists) {
        if (!designFolder.create()) {
            alert("Could not create the Design subfolder at:\n" + designFolder.fsName);
            return;
        }
    }

    var overallStartTime = new Date().getTime();
    var successCount  = 0;
    var convertedCount = 0;
    var errorCount    = 0;
    var errorMessages = [];
    var savedPdfPaths = [];

    // Snapshot doc names up front (before BT might poison the main bridge)
    var docNames = [];
    for (var i = 0; i < documents.length; i++) {
        try { docNames.push(documents[i].name); } catch (e) {}
    }

    for (var docIndex = 0; docIndex < docNames.length; docIndex++) {
        var docName = docNames[docIndex];
        var outPath = computePdfOutputPath(docName, designFolder, prefix);

        // --- BT1: find, activate, convert to CMYK if RGB ---
        var r1 = bridgeTalkExec(buildFindAndConvertScript(docName), 60);
        if (!r1.success || !r1.body || r1.body.indexOf("ok:") !== 0) {
            errorMessages.push(docName + ": convert step - " + (r1.body || "BT failed"));
            errorCount++;
            bridgeTalkExec(buildCloseActiveScript(), 20);
            continue;
        }
        if (r1.body === "ok:CMYK") convertedCount++;
        $.sleep(400); // let the bridge reset between BT messages

        // --- BT2: do wrap-template work, save as PDF, close ---
        var r2 = bridgeTalkExec(buildProcessWrapScript(docName, outPath), 90);
        if (!r2.success || r2.body !== "ok") {
            errorMessages.push(docName + ": process step - " + (r2.body || "BT failed"));
            errorCount++;
            bridgeTalkExec(buildCloseActiveScript(), 20);
            continue;
        }

        successCount++;
        savedPdfPaths.push(outPath);
        $.sleep(300);
    }

    // Optionally reopen the saved PDFs via BT (main bridge may be dead by now)
    if (reopenPdfs) {
        for (var ri = 0; ri < savedPdfPaths.length; ri++) {
            var rOpen = bridgeTalkExec(buildOpenFileScript(savedPdfPaths[ri]), 30);
            if (!rOpen.success || rOpen.body !== "ok") {
                errorMessages.push(savedPdfPaths[ri] + ": reopen failed (" + (rOpen.body || "BT failed") + ")");
            }
            $.sleep(200);
        }
    }

    var totalTime = new Date().getTime() - overallStartTime;

    var resultMessage = "Processing complete!\n\n";
    resultMessage += "Successfully processed: " + successCount + " document(s)\n";
    if (convertedCount > 0) {
        resultMessage += "  (of which " + convertedCount + " were auto-converted from RGB to CMYK)\n";
    }
    resultMessage += "Saved to: " + designFolder.fsName + "\n";
    if (errorCount > 0) {
        resultMessage += "Failed: " + errorCount + " document(s)\n\n";
        if (errorMessages.length > 0) {
            resultMessage += "Issues:\n" + errorMessages.join("\n");
        }
    }
    resultMessage += "\n\nTotal time: " + Math.round(totalTime / 1000) + " seconds";

    alert(resultMessage);
}

// ============================================================================
// COMPUTE PDF OUTPUT PATH (proof-based naming, matches the original)
// e.g. "FooJob_Crew_ShortB_Rear.tif" + prefix "PromaserB" → "PromaserB_Rear.pdf"
// ============================================================================
function computePdfOutputPath(docName, designFolder, prefix) {
    var baseName = docName.replace(/\.[^\/\.]+$/, "");
    var shortName = baseName;
    var lastU = baseName.lastIndexOf("_");
    if (lastU !== -1) shortName = baseName.substring(lastU + 1);
    var finalName = shortName;
    if (prefix && prefix.length > 0) finalName = prefix + "_" + shortName;
    return designFolder.fsName + "/" + finalName + ".pdf";
}

// ============================================================================
// BRIDGE TALK SYNCHRONOUS EXECUTE
// Sends scriptBody to this Illustrator instance and waits for the result by
// spin-sleeping while pumping BridgeTalk. Returns { success, body }.
// ============================================================================
function bridgeTalkExec(scriptBody, timeoutSecs) {
    var done = false;
    var result = null;

    var bt = new BridgeTalk();
    bt.target = BridgeTalk.appSpecifier;
    bt.body = scriptBody;

    bt.onResult = function(resultObj) {
        result = { success: true, body: resultObj.body };
        done = true;
    };
    bt.onError = function(errObj) {
        result = { success: false, body: (errObj && errObj.body) ? errObj.body : "BridgeTalk error" };
        done = true;
    };
    bt.onTimeout = function() {
        result = { success: false, body: "BridgeTalk timeout" };
        done = true;
    };

    try { bt.send(timeoutSecs); } catch (eSend) {
        return { success: false, body: "BridgeTalk send failed: " + eSend.message };
    }

    var maxWaitMs = (timeoutSecs + 10) * 1000;
    var waited = 0;
    while (!done && waited < maxWaitMs) {
        $.sleep(200);
        waited += 200;
        try { BridgeTalk.pump(); } catch (ePump) {}
    }

    return result || { success: false, body: "BridgeTalk wait timeout" };
}

// ============================================================================
// BUILD FIND-AND-CONVERT SCRIPT (BridgeTalk message 1 per file)
// Finds an open doc by name, activates it, runs the CMYK menu command if RGB.
// Returns "ok:CMYK" (converted), "ok:RGB" (already CMYK), or "error:<msg>".
// ============================================================================
function buildFindAndConvertScript(docName) {
    var nameEsc = docName.replace(/\\/g, "\\\\").replace(/'/g, "\\'");
    return (
        "(function() {" +
        "  try {" +
        "    var targetDoc = null;" +
        "    for (var s = 0; s < app.documents.length; s++) {" +
        "      try {" +
        "        if (app.documents[s].name === '" + nameEsc + "') {" +
        "          targetDoc = app.documents[s];" +
        "          break;" +
        "        }" +
        "      } catch (eF) {}" +
        "    }" +
        "    if (!targetDoc) return 'error:doc not found in open documents';" +
        "    app.activeDocument = targetDoc;" +
        "    $.sleep(250);" +
        "    app.redraw();" +
        "    $.sleep(150);" +
        "    var d = app.activeDocument;" +
        "    var cs = d.documentColorSpace;" +
        "    if (cs === DocumentColorSpace.RGB) {" +
        // Menu command does the conversion but throws a spurious error.
        // Swallow it; the next BT message (a fresh context) will pick up
        // the now-CMYK active doc.
        "      try { app.executeMenuCommand('doc-color-cmyk'); } catch (eMenu) {}" +
        "      $.sleep(1500);" +
        "      return 'ok:CMYK';" +
        "    } else {" +
        "      return 'ok:RGB';" + // already CMYK upstream
        "    }" +
        "  } catch (eOuter) {" +
        "    return 'error:' + eOuter.message;" +
        "  }" +
        "})();"
    );
}

// ============================================================================
// BUILD PROCESS-WRAP SCRIPT (BridgeTalk message 2 per file)
// Finds the doc by name, performs all wrap-template operations, saves as PDF.
// Runs in a fresh BT context so the bridge is healthy again after BT1's
// menu command broke it.
//
// Operations (match the legacy processWrapTemplate logic):
//   1. Delete top layer if it's a "Your Logo Here" / "Logo" layer
//   2. Find "Design on Layers Below" layer, clear contents, rename to "Design"
//   3. Rename first artboard to "Proof"
//   4. Lock "Template Layers | Do not edit!" layer
//   5. Set active layer to "Design"
//   6. SaveAs PDF using [Illustrator Default] preset
//   7. Close doc
// ============================================================================
function buildProcessWrapScript(docName, outputPath) {
    var nameEsc = docName.replace(/\\/g, "\\\\").replace(/'/g, "\\'");
    var pathEsc = outputPath.replace(/\\/g, "\\\\").replace(/'/g, "\\'");
    return (
        "(function() {" +
        "  function _clearLayerContents(layer) {" +
        "    try {" +
        "      if (layer.layers && layer.layers.length > 0) {" +
        "        for (var ii = layer.layers.length - 1; ii >= 0; ii--) layer.layers[ii].remove();" +
        "      }" +
        "      if (layer.pageItems && layer.pageItems.length > 0) {" +
        "        for (var jj = layer.pageItems.length - 1; jj >= 0; jj--) layer.pageItems[jj].remove();" +
        "      }" +
        "    } catch (e) {}" +
        "  }" +
        "  function _findLayerByName(layersCollection, targetName) {" +
        "    for (var ii = 0; ii < layersCollection.length; ii++) {" +
        "      var layer = layersCollection[ii];" +
        "      if (layer.name === targetName) return layer;" +
        "      if (layer.layers && layer.layers.length > 0) {" +
        "        var found = _findLayerByName(layer.layers, targetName);" +
        "        if (found) return found;" +
        "      }" +
        "    }" +
        "    return null;" +
        "  }" +
        "  try {" +
        "    var targetDoc = null;" +
        "    for (var s = 0; s < app.documents.length; s++) {" +
        "      try {" +
        "        if (app.documents[s].name === '" + nameEsc + "') {" +
        "          targetDoc = app.documents[s];" +
        "          break;" +
        "        }" +
        "      } catch (eF) {}" +
        "    }" +
        "    if (!targetDoc) return 'error:doc not found after convert';" +
        "    app.activeDocument = targetDoc;" +
        "    $.sleep(200);" +
        "    var doc = app.activeDocument;" +
        // 1. Delete logo top layer
        "    if (doc.layers.length > 0) {" +
        "      var topLayer = doc.layers[0];" +
        "      if (topLayer.name.indexOf('Your Logo Here') !== -1 || topLayer.name.indexOf('Logo') !== -1) {" +
        "        topLayer.remove();" +
        "      }" +
        "    }" +
        // 2. Find & clear "Design on Layers Below", rename to "Design"
        "    var designLayer = null;" +
        "    for (var i = 0; i < doc.layers.length; i++) {" +
        "      if (doc.layers[i].name.indexOf('Design on Layers Below') !== -1) {" +
        "        designLayer = doc.layers[i];" +
        "        break;" +
        "      }" +
        "    }" +
        "    if (designLayer) {" +
        "      _clearLayerContents(designLayer);" +
        "      designLayer.name = 'Design';" +
        "    }" +
        // 3. Rename artboard
        "    if (doc.artboards.length > 0) doc.artboards[0].name = 'Proof';" +
        // 4. Lock template layer
        "    var templateLayer = _findLayerByName(doc.layers, 'Template Layers | Do not edit!');" +
        "    if (templateLayer) templateLayer.locked = true;" +
        // 5. Set active layer to "Design"
        "    var newDesignLayer = _findLayerByName(doc.layers, 'Design');" +
        "    if (newDesignLayer) doc.activeLayer = newDesignLayer;" +
        // 6. Save as PDF
        "    var outFile = new File('" + pathEsc + "');" +
        "    var pdfOpts = new PDFSaveOptions();" +
        "    pdfOpts.pdfPreset = '[Illustrator Default]';" +
        "    doc.saveAs(outFile, pdfOpts);" +
        "    $.sleep(300);" +
        // 7. Close
        "    try { doc.close(SaveOptions.DONOTSAVECHANGES); } catch (eClose) {}" +
        "    return 'ok';" +
        "  } catch (eOuter) {" +
        "    return 'error:' + eOuter.message;" +
        "  }" +
        "})();"
    );
}

// ============================================================================
// BUILD OPEN-FILE SCRIPT (used for the optional reopen at the end)
// ============================================================================
function buildOpenFileScript(filePath) {
    var pEsc = filePath.replace(/\\/g, "\\\\").replace(/'/g, "\\'");
    return (
        "(function() {" +
        "  try {" +
        "    var f = new File('" + pEsc + "');" +
        "    if (!f.exists) return 'error:file not found';" +
        "    app.open(f);" +
        "    $.sleep(300);" +
        "    return 'ok';" +
        "  } catch (e) {" +
        "    return 'error:' + e.message;" +
        "  }" +
        "})();"
    );
}

// ============================================================================
// BUILD CLOSE-ACTIVE SCRIPT (cleanup if anything fails mid-iteration)
// ============================================================================
function buildCloseActiveScript() {
    return (
        "(function() {" +
        "  try {" +
        "    if (app.documents.length > 0) {" +
        "      app.activeDocument.close(SaveOptions.DONOTSAVECHANGES);" +
        "    }" +
        "    return 'ok';" +
        "  } catch (e) {" +
        "    return 'error:' + e.message;" +
        "  }" +
        "})();"
    );
}

// ============================================================================
// LEGACY FUNCTIONS (no longer called by the BT-based convert flow, but kept
// in case other tooling or future scripts call into them)
// ============================================================================
// NOTE: BridgeTalk script bodies (buildProcessWrapScript above) inline their
// own copies of clearLayerContents/findLayerByName because BT scripts run in
// a fresh execution context that can't reference outer-scope functions.
// These outer-scope copies remain here as reusable utilities.
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
