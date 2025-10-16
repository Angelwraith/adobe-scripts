/*
@METADATA
{
  "name": "New Proof From Template",
  "description": "Create a blank proof from a Sheet file",
  "version": "1.2",
  "target": "illustrator",
  "tags": ["Optimal", "Prime", "processors", "scaleFactor"]
}
@END_METADATA
*/


// =============================================================================
// CONFIGURATION
// =============================================================================
// Automatically detect OS and set appropriate template path
var TEMPLATE_PATH = ($.os.indexOf("Windows") !== -1) 
    ? "D:/COMMON TEMPLATES/1_PROOF_60pg.ait"  // Windows path
    : "/Users/Jenny/Desktop/DPS/_Templates/Proofs/1_PROOF_60pg.ait";  // Mac path

// =============================================================================
// MAIN SCRIPT
// =============================================================================

prepareTemplate();

function prepareTemplate() {
    try {
        // Show input dialog
        var dialogResult = showInputDialog();
        if (!dialogResult) {
            return; // User cancelled
        }
        
        var projectFolder = dialogResult.projectFolder;
        var lineItemCount = dialogResult.lineItemCount;
        
        // Validate inputs
        validateInputs(projectFolder, lineItemCount);
        
        // Open template
        var templateFile = File(TEMPLATE_PATH);
        if (!templateFile.exists) {
            throw new Error('Template file not found at: ' + TEMPLATE_PATH);
        }
        
        // Suppress all dialogs during file opening
        var originalInteractionLevel = app.userInteractionLevel;
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;
        
        var doc;
        try {
            doc = app.open(templateFile);
        } finally {
            // Restore original interaction level
            app.userInteractionLevel = originalInteractionLevel;
        }
        
        // Validate template structure
        validateTemplate(doc);
        
        // Find and relink files
        var sheet1File = findSheet1File(projectFolder);
        relinkAllItems(doc, sheet1File);
        
        // Delete excess artboards (including Details 2 if not needed)
        deleteExcessArtboards(doc, lineItemCount);
        
        // Lock "Locked Notes" layer
        lockNotesLayer(doc);
        
        // Unlock "Unlocked Notes" layer before saving
        unlockNotesLayer(doc);
        
        // Save file with new name
        saveTemplateFile(doc, sheet1File, projectFolder);
        
    } catch (e) {
        alert('Error: ' + e.message);
    }
}

function showInputDialog() {
    var dialog = new Window("dialog", "Template Preparation");
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.spacing = 10;
    dialog.margins = 20;
    
    // Project folder section
    var folderGroup = dialog.add("group");
    folderGroup.orientation = "column";
    folderGroup.alignChildren = "fill";
    
    folderGroup.add("statictext", undefined, "Project Folder:");
    var folderPanel = folderGroup.add("panel");
    folderPanel.orientation = "row";
    folderPanel.alignChildren = "fill";
    folderPanel.margins = 10;
    
    var folderInput = folderPanel.add("edittext", undefined, "");
    folderInput.characters = 40;
    
    var browseButton = folderPanel.add("button", undefined, "Browse...");
    browseButton.onClick = function() {
        var folder = Folder.selectDialog("Select project folder:");
        if (folder) {
            folderInput.text = folder.fsName;
        }
    };
    
    // Line items section
    var itemsGroup = dialog.add("group");
    itemsGroup.orientation = "column";
    itemsGroup.alignChildren = "fill";
    
    itemsGroup.add("statictext", undefined, "Number of Line Items (1-60):");
    var itemsPanel = itemsGroup.add("panel");
    itemsPanel.orientation = "row";
    itemsPanel.alignChildren = "center";
    itemsPanel.margins = 10;
    
    var itemsInput = itemsPanel.add("edittext", undefined, "1");
    itemsInput.characters = 5;
    itemsInput.justify = "center";
    
    var buttonGroup = itemsPanel.add("group");
    buttonGroup.orientation = "column";
    buttonGroup.spacing = 2;
    
    var upButton = buttonGroup.add("button", undefined, "+");
    upButton.preferredSize.width = 25;
    upButton.preferredSize.height = 20;
    
    var downButton = buttonGroup.add("button", undefined, "-");
    downButton.preferredSize.width = 25;
    downButton.preferredSize.height = 20;
    
    // Input validation and formatting
    itemsInput.onChange = function() {
        var value = parseInt(this.text);
        if (isNaN(value) || value < 1) {
            this.text = "1";
        } else if (value > 60) {
            this.text = "60";
        }
    };
    
    // Button functionality
    upButton.onClick = function() {
        var current = parseInt(itemsInput.text) || 1;
        var increment = ScriptUI.environment.keyboardState.shiftKey ? 10 : 1;
        var newValue = Math.min(60, current + increment);
        itemsInput.text = newValue.toString();
    };
    
    downButton.onClick = function() {
        var current = parseInt(itemsInput.text) || 1;
        var decrement = ScriptUI.environment.keyboardState.shiftKey ? 10 : 1;
        var newValue = Math.max(1, current - decrement);
        itemsInput.text = newValue.toString();
    };
    
    // Help text
    var helpText = dialog.add("statictext", undefined, "Tip: Hold Shift while clicking arrows to change by 10");
    helpText.graphics.foregroundColor = helpText.graphics.newPen(helpText.graphics.PenType.SOLID_COLOR, [0.5, 0.5, 0.5], 1);
    
    // Buttons
    var buttonGroup2 = dialog.add("group");
    buttonGroup2.alignment = "center";
    
    var cancelButton = buttonGroup2.add("button", undefined, "Cancel");
    var okButton = buttonGroup2.add("button", undefined, "Continue");
    
    cancelButton.onClick = function() {
        dialog.close(0);
    };
    
    okButton.onClick = function() {
        dialog.close(1);
    };
    
    // Show dialog
    var result = dialog.show();
    
    if (result === 1) {
        return {
            projectFolder: folderInput.text,
            lineItemCount: parseInt(itemsInput.text) || 1
        };
    }
    
    return null;
}

function validateInputs(projectFolder, lineItemCount) {
    // Check if project folder exists
    var folder = Folder(projectFolder);
    if (!folder.exists) {
        throw new Error('Project folder does not exist: ' + projectFolder);
    }
    
    // Validate line item count
    if (isNaN(lineItemCount) || lineItemCount < 1 || lineItemCount > 60) {
        throw new Error('Line item count must be between 1 and 60');
    }
}

function validateTemplate(doc) {
    // Check if template has expected number of artboards (62 total: 1 details + 1 details 2 + 60 line items)
    if (doc.artboards.length !== 62) {
        throw new Error('Template should have exactly 62 artboards (1 Details + 1 Details 2 + 60 line items). Found: ' + doc.artboards.length);
    }
    
    // Check for "Locked Notes" layer
    var hasLockedNotesLayer = false;
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === "Locked Notes") {
            hasLockedNotesLayer = true;
            break;
        }
    }
    
    if (!hasLockedNotesLayer) {
        throw new Error('Template must contain a layer named "Locked Notes"');
    }
}

function findSheet1File(projectFolder) {
    var folder = Folder(projectFolder);
    var files = folder.getFiles("*");
    var sheet1Files = [];
    
    for (var i = 0; i < files.length; i++) {
        if (files[i] instanceof File) {
            var fileName = files[i].name;
            var decodedFileName = decodeURI(fileName);
            
            // Check if decoded filename ends with "- Sheet1" followed by any extension
            if (decodedFileName.match(/\s*-\s*Sheet1\.[^.]*$/i)) {
                sheet1Files.push(files[i]);
            }
        }
    }
    
    if (sheet1Files.length === 0) {
        // Debug: List all files in folder for troubleshooting (decoded names)
        var allFiles = [];
        for (var j = 0; j < files.length; j++) {
            if (files[j] instanceof File) {
                allFiles.push(decodeURI(files[j].name));
            }
        }
        throw new Error('No file ending with "- Sheet1.[extension]" found in project folder.\nFiles found: ' + allFiles.join(', '));
    }
    
    if (sheet1Files.length > 1) {
        throw new Error('Multiple files ending with "- Sheet1" found. Please ensure only one exists.');
    }
    
    return sheet1Files[0];
}

function relinkAllItems(doc, newLinkFile) {
    // Simple approach: relink all items (protected items should be on locked layers)
    var placed = [];
    var rasters = [];
    
    // Get all placed items
    for (var i = 0; i < doc.placedItems.length; i++) {
        placed.push(doc.placedItems[i]);
    }
    
    // Get all raster items
    for (var i = 0; i < doc.rasterItems.length; i++) {
        rasters.push(doc.rasterItems[i]);
    }
    
    if (placed.length === 0 && rasters.length === 0) {
        return; // No items to relink
    }
    
    // Create action for raster relinking if needed
    var act = null;
    if (rasters.length > 0) {
        var actName = 'template_relink_' + ((+new Date()) * Math.random() * 1e7 + new Array(6).join('0')).slice(0, 6);
        var actionString = getFastRelinkActString(newLinkFile.fsName, actName);
        act = new Action(actName, actName, actionString);
        act.loadAction();
    }
    
    try {
        // Relink placed items
        relinkAllPlaced(placed, newLinkFile);
        
        // Relink raster items
        if (rasters.length > 0) {
            relinkAllRasters(rasters, act);
        }
        
    } finally {
        // Clean up action
        if (act) {
            try {
                act.rmAction();
            } catch (e) {
                // Ignore cleanup errors
            }
        }
    }
}

function relinkAllPlaced(placed, newLinkFile) {
    for (var i = 0; i < placed.length; i++) {
        try {
            placed[i].file = newLinkFile;
        } catch (e) {
            // Skip items that can't be relinked (locked items, embedded items, etc.)
        }
    }
    app.executeMenuCommand('deselectall');
}

function relinkAllRasters(rasters, act) {
    for (var i = rasters.length - 1; i >= 0; i--) {
        try {
            rasters[i].selected = true;
            act.runAction();
            app.executeMenuCommand('deselectall');
        } catch (e) {
            // Skip items that can't be relinked (locked items, embedded items, etc.)
            app.executeMenuCommand('deselectall');
        }
    }
}

function resizeDetailMasks(doc, lineItemCount) {
    // Find "Locked Notes" layer
    var lockedNotesLayer = null;
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === "Locked Notes") {
            lockedNotesLayer = doc.layers[i];
            break;
        }
    }
    
    if (!lockedNotesLayer) {
        throw new Error('Could not find "Locked Notes" layer for mask resizing');
    }
    
    // Ensure layer is unlocked for editing
    lockedNotesLayer.locked = false;
    
    // Calculate mask sizes
    var mask1Items = Math.min(lineItemCount, 32);
    var mask2Items = Math.max(0, lineItemCount - 32);
    
    // Details Mask 1 sizing: 8.7415" for 1 item, 0.8675" for 32 items
    var mask1Height = 8.7415 - ((mask1Items - 1) * 0.2538);
    
    // Details Mask 2 sizing: 8.7415" for 1 item (line 33), 1.888" for 28 items (lines 33-60)
    var mask2Height = 0;
    if (mask2Items > 0) {
        mask2Height = 8.7415 - ((mask2Items - 1) * 0.2538);
    }
    
    // Resize Details Mask 1
    resizeMaskObject(lockedNotesLayer, "Details Mask 1", mask1Height);
    
    // Resize Details Mask 2 if needed
    if (lineItemCount > 32) {
        resizeMaskObject(lockedNotesLayer, "Details Mask 2", mask2Height);
    }
}

function resizeMaskObject(layer, maskName, newHeight) {
    // Find the mask object in the layer
    var maskObject = findObjectInLayer(layer, maskName);
    
    if (!maskObject) {
        throw new Error('Could not find "' + maskName + '" in Locked Notes layer');
    }
    
    // Ensure object is unlocked
    maskObject.locked = false;
    
    // Get current bounds [left, top, right, bottom]
    var bounds = maskObject.geometricBounds;
    var newHeightPoints = newHeight * 72; // Convert inches to points
    
    // Calculate new bounds anchored at bottom
    var newBounds = [
        bounds[0],                              // left (unchanged)
        bounds[3] + newHeightPoints,           // top (bottom + new height)
        bounds[2],                              // right (unchanged)
        bounds[3]                               // bottom (unchanged - anchored)
    ];
    
    // Apply new bounds
    maskObject.geometricBounds = newBounds;
}

function getLayerObjectNames(layer) {
    var names = [];
    for (var i = 0; i < layer.pageItems.length; i++) {
        names.push(layer.pageItems[i].name || 'unnamed object');
    }
    return names.join(', ');
}

function findObjectInLayer(layer, objectName) {
    // Search through page items in the layer
    for (var i = 0; i < layer.pageItems.length; i++) {
        if (layer.pageItems[i].name === objectName) {
            return layer.pageItems[i];
        }
    }
    
    // Search through sublayers recursively
    for (var j = 0; j < layer.layers.length; j++) {
        var found = findObjectInLayer(layer.layers[j], objectName);
        if (found) {
            return found;
        }
    }
    
    return null;
}

function deleteExcessArtboards(doc, lineItemCount) {
    var totalArtboards = doc.artboards.length;
    
    // Calculate which artboards to keep
    var artboardsToKeep = [];
    if (lineItemCount <= 32) {
        // Keep: Details (0) + line items (2 to lineItemCount+1), skip Details 2 (1)
        artboardsToKeep.push(0); // Details page
        // Add line item pages (start from index 2, skip Details 2)
        for (var i = 2; i < 2 + lineItemCount; i++) {
            artboardsToKeep.push(i);
        }
    } else {
        // Keep: Details (0) + Details 2 (1) + line items (2 to lineItemCount+1)
        artboardsToKeep = [0, 1]; // Details and Details 2 pages
        // Add all line item pages
        for (var i = 2; i < 2 + lineItemCount; i++) {
            artboardsToKeep.push(i);
        }
    }
    
    // PHASE 1: Lock content on artboards we want to keep
    app.executeMenuCommand('deselectall');
    
    for (var i = 0; i < artboardsToKeep.length; i++) {
        var artboardIndex = artboardsToKeep[i];
        doc.artboards.setActiveArtboardIndex(artboardIndex);
        app.executeMenuCommand('selectallinartboard');
        app.executeMenuCommand('lock');
    }
    
    // PHASE 2: Select and delete all remaining unlocked content in one operation
    app.executeMenuCommand('selectall');
    app.executeMenuCommand('clear'); // Delete without using clipboard
    
    // PHASE 3: Unlock all content
    app.executeMenuCommand('unlockAll');
    
    // PHASE 4: Delete excess artboards in reverse order (to maintain indices)
    var artboardsToDelete = [];
    for (var i = totalArtboards - 1; i >= 0; i--) {
        var shouldKeep = false;
        for (var j = 0; j < artboardsToKeep.length; j++) {
            if (i === artboardsToKeep[j]) {
                shouldKeep = true;
                break;
            }
        }
        if (!shouldKeep) {
            artboardsToDelete.push(i);
        }
    }
    
    // Delete artboards in batch (they should be empty now)
    for (var i = 0; i < artboardsToDelete.length; i++) {
        var artboardIndex = artboardsToDelete[i];
        doc.artboards.remove(artboardIndex);
    }
    
    // Set active artboard back to the first one
    if (doc.artboards.length > 0) {
        doc.artboards.setActiveArtboardIndex(0);
    }
    
    app.executeMenuCommand('deselectall');
}

function lockNotesLayer(doc) {
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === "Locked Notes") {
            doc.layers[i].locked = true;
            break;
        }
    }
}

function unlockNotesLayer(doc) {
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === "Unlocked Notes") {
            doc.layers[i].locked = false;
            break;
        }
    }
}

function saveTemplateFile(doc, sheet1File, projectFolder) {
    // Extract the base name and replace "- Sheet1.[extension]" with "_Proof"
    var originalName = decodeURI(sheet1File.name);
    var nameWithoutExtension = originalName.replace(/\.[^\.]+$/, ''); // Remove file extension
    var newName = nameWithoutExtension.replace(/\s*-\s*Sheet1$/i, "_Proof");
    
    var saveFile = new File(projectFolder + "/" + newName + ".pdf");
    
    // Save as PDF with specified settings
    var pdfOptions = new PDFSaveOptions();
    pdfOptions.pdfPreset = "[Illustrator Default]";
    pdfOptions.acrobatLayers = true;
    pdfOptions.colorCompression = CompressionQuality.None;
    pdfOptions.compatibility = PDFCompatibility.ACROBAT7; // PDF 1.6
    
    doc.saveAs(saveFile, pdfOptions);
}

// =============================================================================
// UTILITY CLASSES AND FUNCTIONS (from your Fast_Relink script)
// =============================================================================

function Action(actionName, setName, actionString) {
    this.file = new File('~/JavaScriptAction.aia');
    this.loadAction = function () {
        this.file.open('w');
        this.file.write(actionString);
        this.file.close();
        app.loadAction(this.file);
    };
    this.runAction = function () {
        app.doScript(actionName, setName);
    };
    this.rmAction = function () {
        app.unloadAction(setName, '');
        this.file.remove();
    };
}

function getFastRelinkActString(newLinkFileFsPath, actName) {
    var actNameEncoded = encodeStr2Ansii(actName);
    var newLinkFileFsPathEncoded = encodeStr2Ansii(newLinkFileFsPath);
    
    var actRelinkString = "/version 3" +
        "/name [ " +
        actNameEncoded.length / 2 + " " +
        actNameEncoded +
        "]" +
        "/isOpen 0" +
        "/actionCount 1" +
        "/action-1 {" +
        "	/name [ " +
        actNameEncoded.length / 2 + " " +
        actNameEncoded +
        "	]" +
        "	/keyIndex 0" +
        "	/colorIndex 0" +
        "	/isOpen 1" +
        "	/eventCount 1" +
        "	/event-1 {" +
        "		/useRulersIn1stQuadrant 0" +
        "		/internalName (adobe_placeDocument)" +
        "		/localizedName [ 5" +
        "			506c616365" +
        "		]" +
        "		/isOpen 0" +
        "		/isOn 1" +
        "		/hasDialog 1" +
        "		/showDialog 0" +
        "		/parameterCount 12" +
        "		/parameter-1 {" +
        "			/key 1885431653" +
        "			/showInPalette -1" +
        "			/type (integer)" +
        "			/value 1" +
        "		}" +
        "		/parameter-2 {" +
        "			/key 1668444016" +
        "			/showInPalette -1" +
        "			/type (enumerated)" +
        "			/name [ 7" +
        "				43726f7020546f" +
        "			]" +
        "			/value 1" +
        "		}" +
        "		/parameter-3 {" +
        "			/key 1885823860" +
        "			/showInPalette -1" +
        "			/type (integer)" +
        "			/value 1" +
        "		}" +
        "		/parameter-4 {" +
        "			/key 1851878757" +
        "			/showInPalette -1" +
        "			/type (ustring)" +
        "			/value [ " +
        newLinkFileFsPathEncoded.length / 2 + " " +
        newLinkFileFsPathEncoded +
        "			]" +
        "		}" +
        "		/parameter-5 {" +
        "			/key 1818848875" +
        "			/showInPalette -1" +
        "			/type (boolean)" +
        "			/value 1" +
        "		}" +
        "		/parameter-6 {" +
        "			/key 1919970403" +
        "			/showInPalette -1" +
        "			/type (boolean)" +
        "			/value 1" +
        "		}" +
        "		/parameter-7 {" +
        "			/key 1953329260" +
        "			/showInPalette -1" +
        "			/type (boolean)" +
        "			/value 0" +
        "		}" +
        "		/parameter-8 {" +
        "			/key 1768779887" +
        "			/showInPalette -1" +
        "			/type (boolean)" +
        "			/value 0" +
        "		}" +
        "		/parameter-9 {" +
        "			/key 1885828462" +
        "			/showInPalette -1" +
        "			/type (boolean)" +
        "			/value 0" +
        "		}" +
        "		/parameter-10 {" +
        "			/key 1935895653" +
        "			/showInPalette -1" +
        "			/type (real)" +
        "			/value 1.0" +
        "		}" +
        "		/parameter-11 {" +
        "			/key 1953656440" +
        "			/showInPalette -1" +
        "			/type (real)" +
        "			/value 0.0" +
        "		}" +
        "		/parameter-12 {" +
        "			/key 1953656441" +
        "			/showInPalette -1" +
        "			/type (real)" +
        "			/value 0.0" +
        "		}" +
        "	}" +
        "}";

    return actRelinkString;
}

function encodeStr2Ansii(str) {
    var result = '';
    for (var i = 0; i < str.length; i++) {
        var chr = File.encode(str[i]);
        chr = chr.replace(/%/gmi, '');
        if (chr.length == 1) {
            result += chr.charCodeAt(0).toString(16);
        } else {
            result += chr.toLowerCase();
        }
    }
    return result;
}