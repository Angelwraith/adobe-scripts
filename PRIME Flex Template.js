#target illustrator

/*@METADATA{
  "name": "PRIME Flex Template",
  "description": "Create artboards with reg dots for flexible materials",
  "version": "3.5",
  "target": "illustrator",
  "tags": ["artboard", "template", "setup"]
}@END_METADATA*/

// Main entry point
try {
    var docChoice;
    if (app.documents.length === 0) {
        // Need PRIME folder for new document
        docChoice = promptForPrimeFolderOnly();
    } else {
        docChoice = showDocumentChoiceDialog();
    }
    
    if (docChoice) {
        showSetupDialog(docChoice);
    }
} catch (e) {
    alert("Startup error: " + e.toString());
}

function promptForPrimeFolderOnly() {
    var dialog = new Window("dialog", "PRIME Flex Template");
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.spacing = 15;
    dialog.margins = 30;
    
    dialog.add("statictext", undefined, "Create new PRIME file in PRIME folder:");
    var primeFolderField = dialog.add("edittext", undefined, "");
    primeFolderField.preferredSize.width = 400;
    
    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.spacing = 10;
    buttonGroup.alignment = "center";
    
    var cancelBtn = buttonGroup.add("button", undefined, "Cancel");
    var continueBtn = buttonGroup.add("button", undefined, "New");
    continueBtn.preferredSize.width = 100;
    
    cancelBtn.onClick = function() { dialog.close(0); };
    continueBtn.onClick = function() {
        if (primeFolderField.text === "") {
            alert("Please enter the PRIME folder path.");
            return;
        }
        dialog.close(1);
    };
    
    if (dialog.show() === 1) {
        return {
            mode: "new",
            primeFolder: primeFolderField.text
        };
    }
    return null;
}

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

function showDocumentChoiceDialog() {
    // Find PRIME documents
    var primeDocuments = [];
    for (var i = 0; i < app.documents.length; i++) {
        var docName = app.documents[i].name;
        if (docName.match(/PRIME/i)) {
            primeDocuments.push({
                doc: app.documents[i],
                name: docName
            });
        }
    }
    
    // Find Proof documents
    var proofDocuments = findProofDocuments();
    
    var dialog = new Window("dialog", "PRIME Flex Template - Document Setup");
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.spacing = 15;
    dialog.margins = 30;
    
    var messageText = dialog.add("statictext", undefined, "Choose document option:", {multiline: false});
    messageText.graphics.font = ScriptUI.newFont("dialog", "Bold", 12);
    
    dialog.add("panel");
    
    // OPTION 1: Add to existing PRIME document
    var docList = null;
    var addDocBtn = null;
    
    if (primeDocuments.length > 0) {
        var addLabel = dialog.add("statictext", undefined, "Add to existing PRIME document:");
        addLabel.graphics.font = ScriptUI.newFont("dialog", "Bold", 11);
        
        var docListGroup = dialog.add("group");
        docListGroup.orientation = "row";
        docListGroup.spacing = 10;
        
        docList = docListGroup.add("dropdownlist", undefined, []);
        docList.preferredSize.width = 400;
        
        for (var i = 0; i < primeDocuments.length; i++) {
            docList.add("item", primeDocuments[i].name);
        }
        docList.selection = 0;
        
        addDocBtn = docListGroup.add("button", undefined, "Add");
        addDocBtn.preferredSize.width = 80;
        
        dialog.add("panel");
    }
    
    // OPTION 2: Create new from Proof document
    var proofList = null;
    var newFromProofBtn = null;
    
    if (proofDocuments.length > 0) {
        var proofLabel = dialog.add("statictext", undefined, "Create new PRIME from open Proof document:");
        proofLabel.graphics.font = ScriptUI.newFont("dialog", "Bold", 11);
        
        var proofListGroup = dialog.add("group");
        proofListGroup.orientation = "row";
        proofListGroup.spacing = 10;
        
        proofList = proofListGroup.add("dropdownlist", undefined, []);
        proofList.preferredSize.width = 400;
        
        for (var i = 0; i < proofDocuments.length; i++) {
            proofList.add("item", proofDocuments[i].name);
        }
        proofList.selection = 0;
        
        newFromProofBtn = proofListGroup.add("button", undefined, "New");
        newFromProofBtn.preferredSize.width = 80;
        
        dialog.add("panel");
    }
    
    // OPTION 3: Create new with manual folder path
    var folderLabel = dialog.add("statictext", undefined, "Create new PRIME file (enter folder path manually):");
    folderLabel.graphics.font = ScriptUI.newFont("dialog", "Bold", 11);
    
    var folderGroup = dialog.add("group");
    folderGroup.orientation = "row";
    folderGroup.spacing = 10;
    
    var primeFolderField = folderGroup.add("edittext", undefined, "");
    primeFolderField.preferredSize.width = 400;
    primeFolderField.helpTip = "Paste the path to the PRIME folder";
    
    var newDocBtn = folderGroup.add("button", undefined, "New");
    newDocBtn.preferredSize.width = 80;
    
    dialog.add("panel");
    
    var cancelBtn = dialog.add("button", undefined, "Cancel");
    cancelBtn.alignment = "center";
    
    var result = null;
    
    // Button handlers
    if (primeDocuments.length > 0) {
        addDocBtn.onClick = function() {
            var selectedIndex = docList.selection.index;
            result = {
                mode: "add",
                document: primeDocuments[selectedIndex].doc,
                primeFolder: primeFolderField.text
            };
            dialog.close();
        };
    }
    
    if (proofDocuments.length > 0) {
        newFromProofBtn.onClick = function() {
            var selectedIndex = proofList.selection.index;
            var selectedDoc = proofDocuments[selectedIndex].document;
            
            // Get the document's folder path
            try {
                var docFile = selectedDoc.fullName;
                var docFolder = docFile.parent;
                result = {
                    mode: "new",
                    primeFolder: docFolder.fsName
                };
                dialog.close();
            } catch (e) {
                alert("Could not determine folder path for this document.\nThe document may not have been saved yet.\nPlease save it first or use the manual folder path option.");
            }
        };
    }
    
    newDocBtn.onClick = function() {
        if (primeFolderField.text === "") {
            alert("Please enter the PRIME folder path.");
            return;
        }
        result = {
            mode: "new",
            primeFolder: primeFolderField.text
        };
        dialog.close();
    };
    
    cancelBtn.onClick = function() {
        result = null;
        dialog.close();
    };
    
    dialog.show();
    return result;
}

function parseExistingArtboards(doc) {
    var materials = {};
    
    // Get scale factor
    var scale = 1;
    try {
        scale = doc.scaleFactor;
    } catch (e) {
        scale = 1;
    }
    
    for (var i = 0; i < doc.artboards.length; i++) {
        var artboard = doc.artboards[i];
        var artboardName = artboard.name;
        
        // Parse names like "3mmACM_Pt1", "3mmACM_Pt2" or "3mmACM" (single)
        var match = artboardName.match(/^(.+?)(?:_Pt(\d+))?$/);
        
        if (match) {
            var materialName = match[1];
            var partNum = match[2] ? parseInt(match[2]) : 1;
            
            // Get artboard dimensions - apply scale factor correction
            var bounds = artboard.artboardRect;
            var widthPoints = bounds[2] - bounds[0];
            var heightPoints = bounds[1] - bounds[3];
            
            // Convert to inches with scale factor correction
            var width = Math.round((widthPoints * scale) / 72);
            var height = Math.round((heightPoints * scale) / 72);
            
            var sizeKey = width + "x" + height;
            
            if (!materials[materialName]) {
                materials[materialName] = {
                    name: materialName,
                    width: width,
                    height: height,
                    maxPart: partNum,
                    sizeKey: sizeKey
                };
            } else {
                if (partNum > materials[materialName].maxPart) {
                    materials[materialName].maxPart = partNum;
                }
            }
        }
    }
    
    var result = [];
    for (var key in materials) {
        result.push(materials[key]);
    }
    
    return result;
}

function showSetupDialog(docChoice) {
    var existingMaterials = [];
    if (docChoice.mode === "add") {
        existingMaterials = parseExistingArtboards(docChoice.document);
    }
    
    var dialog = new Window("dialog", "PRIME Flex Template Setup");
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.spacing = 15;
    dialog.margins = 20;
    dialog.preferredSize.width = 750;
    
    var canvasDropdown = null;
    if (docChoice.mode === "new") {
        var canvasGroup = dialog.add("group");
        canvasGroup.orientation = "row";
        canvasGroup.add("statictext", undefined, "Canvas Type:");
        canvasDropdown = canvasGroup.add("dropdownlist", undefined, ["Large Canvas", "Standard Canvas"]);
        canvasDropdown.selection = 0;
        canvasDropdown.preferredSize.width = 150;
    }
    
    var regGroup = dialog.add("group");
    regGroup.orientation = "row";
    regGroup.add("statictext", undefined, "Reg Dots:");
    var regDropdown = regGroup.add("dropdownlist", undefined, ["Top/Bottom", "Left/Right", "None"]);
    regDropdown.selection = 1;
    regDropdown.preferredSize.width = 150;
    
    dialog.add("panel");
    
    var tableLabel = dialog.add("statictext", undefined, "Material Specifications:");
    tableLabel.graphics.font = ScriptUI.newFont("dialog", "Bold", 12);
    
    var tablePanel = dialog.add("panel");
    tablePanel.orientation = "column";
    tablePanel.alignChildren = "fill";
    tablePanel.preferredSize.height = 300;
    
    var headerGroup = tablePanel.add("group");
    headerGroup.orientation = "row";
    headerGroup.spacing = 5;
    
    var numHeader = headerGroup.add("statictext", undefined, "#");
    numHeader.preferredSize.width = 30;
    numHeader.graphics.font = ScriptUI.newFont("dialog", "Bold", 10);
    
    var matHeader = headerGroup.add("statictext", undefined, "Material Type");
    matHeader.preferredSize.width = 220;
    matHeader.graphics.font = ScriptUI.newFont("dialog", "Bold", 10);
    
    var sizeHeader = headerGroup.add("statictext", undefined, "Size");
    sizeHeader.preferredSize.width = 200;
    sizeHeader.graphics.font = ScriptUI.newFont("dialog", "Bold", 10);
    
    var qtyHeader = headerGroup.add("statictext", undefined, "Qty");
    qtyHeader.preferredSize.width = 80;
    qtyHeader.graphics.font = ScriptUI.newFont("dialog", "Bold", 10);
    
    tablePanel.add("panel");
    
    var scrollGroup = tablePanel.add("group");
    scrollGroup.orientation = "column";
    scrollGroup.alignChildren = "fill";
    scrollGroup.spacing = 2;
    
    var rows = [];
    
    function createRow(material) {
        var rowGroup = scrollGroup.add("group");
        rowGroup.orientation = "row";
        rowGroup.spacing = 5;
        rowGroup.alignChildren = "top";
        
        var rowNum = rowGroup.add("statictext", undefined, (rows.length + 1).toString());
        rowNum.preferredSize.width = 30;
        rowNum.graphics.font = ScriptUI.newFont("dialog", "Bold", 10);
        
        var materialField = rowGroup.add("edittext", undefined, material ? material.name : "");
        materialField.preferredSize.width = 220;
        
        var sizeContainer = rowGroup.add("group");
        sizeContainer.orientation = "column";
        sizeContainer.alignChildren = "fill";
        sizeContainer.preferredSize.width = 200;
        sizeContainer.spacing = 2;
        
        var sizeDropdown = sizeContainer.add("dropdownlist", undefined, 
            ['48"w x 96"h', '60"w x 120"h', '50"w x 100"h', 'Custom']);
        sizeDropdown.preferredSize.width = 200;
        
        if (material) {
            if (material.width === 48 && material.height === 96) {
                sizeDropdown.selection = 0;
            } else if (material.width === 60 && material.height === 120) {
                sizeDropdown.selection = 1;
            } else if (material.width === 50 && material.height === 100) {
                sizeDropdown.selection = 2;
            } else {
                sizeDropdown.selection = 3;
            }
        }
        
        var customGroup = sizeContainer.add("group");
        customGroup.orientation = "row";
        customGroup.spacing = 5;
        customGroup.visible = material && sizeDropdown.selection && sizeDropdown.selection.index === 3;
        
        customGroup.add("statictext", undefined, "W:");
        var customW = customGroup.add("edittext", undefined, material && sizeDropdown.selection && sizeDropdown.selection.index === 3 ? material.width.toString() : "");
        customW.preferredSize.width = 60;
        
        customGroup.add("statictext", undefined, "H:");
        var customH = customGroup.add("edittext", undefined, material && sizeDropdown.selection && sizeDropdown.selection.index === 3 ? material.height.toString() : "");
        customH.preferredSize.width = 60;
        
        sizeDropdown.onChange = function() {
            customGroup.visible = (sizeDropdown.selection && sizeDropdown.selection.index === 3);
            dialog.layout.layout(true);
        };
        
        var qtyDropdown = rowGroup.add("dropdownlist", undefined, 
            ['1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','Custom']);
        qtyDropdown.selection = material ? Math.min(material.maxPart, 20) - 1 : 0;
        qtyDropdown.preferredSize.width = 80;
        
        var customQtyField = rowGroup.add("edittext", undefined, "");
        customQtyField.preferredSize.width = 60;
        customQtyField.visible = false;
        
        qtyDropdown.onChange = function() {
            customQtyField.visible = (qtyDropdown.selection && qtyDropdown.selection.index === 20);
            dialog.layout.layout(true);
        };
        
        var rowData = {
            group: rowGroup,
            rowNum: rowNum,
            material: materialField,
            sizeDropdown: sizeDropdown,
            customW: customW,
            customH: customH,
            customGroup: customGroup,
            qty: qtyDropdown,
            customQty: customQtyField
        };
        
        rows.push(rowData);
        return rowData;
    }
    
    function updateRowNumbers() {
        for (var i = 0; i < rows.length; i++) {
            rows[i].rowNum.text = (i + 1).toString();
        }
    }
    
    if (existingMaterials.length > 0) {
        for (var i = 0; i < existingMaterials.length; i++) {
            createRow(existingMaterials[i]);
        }
    } else {
        createRow(); // Start with just 1 row
    }
    
    var tableButtonGroup = dialog.add("group");
    tableButtonGroup.alignment = "center";
    tableButtonGroup.spacing = 10;
    
    var addRowButton = tableButtonGroup.add("button", undefined, "Add Row");
    var deleteRowButton = tableButtonGroup.add("button", undefined, "Delete Last Row");
    
    addRowButton.onClick = function() {
        createRow();
        dialog.layout.layout(true);
    };
    
    deleteRowButton.onClick = function() {
        if (rows.length > 1) {
            var lastRow = rows.pop();
            scrollGroup.remove(lastRow.group);
            updateRowNumbers();
            dialog.layout.layout(true);
        } else {
            alert("Must have at least one row.");
        }
    };
    
    dialog.add("panel");
    
    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "center";
    buttonGroup.spacing = 10;
    
    var cancelButton = buttonGroup.add("button", undefined, "Cancel");
    var createButton = buttonGroup.add("button", undefined, docChoice.mode === "add" ? "Add Artboards" : "Create Artboards");
    
    var result = null;
    
    cancelButton.onClick = function() {
        result = null;
        dialog.close();
    };
    
    createButton.onClick = function() {
        var specs = [];
        var hasData = false;
        var usedCombos = {};
        
        for (var i = 0; i < rows.length; i++) {
            var row = rows[i];
            var mat = row.material.text.replace(/^\s+|\s+$/g, '');
            
            if (mat === '') continue;
            hasData = true;
            
            var sizeIndex = row.sizeDropdown.selection ? row.sizeDropdown.selection.index : -1;
            if (sizeIndex < 0) {
                alert("Please select a size for material: " + mat);
                return;
            }
            
            var width, height;
            var sizeKey;
            
            if (sizeIndex === 0) {
                width = 48; height = 96;
                sizeKey = "48x96";
            } else if (sizeIndex === 1) {
                width = 60; height = 120;
                sizeKey = "60x120";
            } else if (sizeIndex === 2) {
                width = 50; height = 100;
                sizeKey = "50x100";
            } else {
                var w = parseFloat(row.customW.text);
                var h = parseFloat(row.customH.text);
                
                if (isNaN(w) || w <= 0 || isNaN(h) || h <= 0) {
                    alert("Please enter valid custom dimensions for: " + mat);
                    return;
                }
                
                width = w;
                height = h;
                sizeKey = "custom_" + w + "x" + h;
            }
            
            var comboKey = mat + "|||" + sizeKey;
            if (usedCombos[comboKey]) {
                alert("Duplicate material and size combination:\n" + mat + " at " + width + "x" + height + "\n\nPlease use unique material+size combinations.");
                return;
            }
            usedCombos[comboKey] = true;
            
            var requestedQty;
            if (row.qty.selection.index === 20) {
                // Custom quantity
                requestedQty = parseInt(row.customQty.text);
                if (isNaN(requestedQty) || requestedQty <= 0) {
                    alert("Please enter a valid custom quantity for: " + mat);
                    return;
                }
            } else {
                requestedQty = parseInt(row.qty.selection.text);
            }
            
            // In add mode, calculate how many NEW artboards to create
            var qtyToCreate = requestedQty;
            var startingPartNumber = 1;
            
            if (docChoice.mode === "add") {
                // Find existing material to see how many already exist
                for (var j = 0; j < existingMaterials.length; j++) {
                    if (existingMaterials[j].name === mat && 
                        existingMaterials[j].width === width && 
                        existingMaterials[j].height === height) {
                        var existingCount = existingMaterials[j].maxPart;
                        qtyToCreate = requestedQty - existingCount;
                        startingPartNumber = existingCount + 1;
                        break;
                    }
                }
                
                if (qtyToCreate <= 0) {
                    continue; // Skip this material - nothing new to add
                }
            }
            
            specs.push({
                material: mat,
                width: width,
                height: height,
                quantity: qtyToCreate,
                startingPartNumber: startingPartNumber
            });
        }
        
        if (!hasData) {
            alert("Please enter at least one material specification.");
            return;
        }
        
        if (docChoice.mode === "add" && specs.length === 0) {
            alert("No new artboards to add. All materials are already at the requested quantity.");
            return;
        }
        
        result = {
            docChoice: docChoice,
            canvasType: canvasDropdown ? (canvasDropdown.selection.index === 0 ? "Large" : "Standard") : null,
            regDots: regDropdown.selection.index === 0 ? "TopBottom" : (regDropdown.selection.index === 1 ? "LeftRight" : "None"),
            specs: specs
        };
        
        dialog.close();
    };
    
    dialog.show();
    
    if (result) {
        createArtboards(result);
    }
}
function calculateSpaceNeeded(specs, scale) {
    var spacing = (10 * 72) / scale;
    var maxSingleArtboardWidth = 0;
    var totalArea = 0;
    var rowCount = 0;
    
    for (var i = 0; i < specs.length; i++) {
        var spec = specs[i];
        var artboardWidth = (spec.width * 72) / scale;
        var artboardHeight = (spec.height * 72) / scale;
        
        // Track the largest single artboard (this must fit)
        if (artboardWidth > maxSingleArtboardWidth) {
            maxSingleArtboardWidth = artboardWidth;
        }
        
        totalArea += (artboardWidth * artboardHeight * spec.quantity);
        rowCount++;
    }
    
    // Only check if a SINGLE artboard fits (wrapping handles multiple)
    return {
        maxRowWidth: maxSingleArtboardWidth,
        totalArea: totalArea,
        estimatedRows: rowCount
    };
}

function createArtboards(config) {
    $.writeln("=== START ARTBOARD CREATION ===");
    
    try {
        var doc;
        var scale = 1;
        
        $.writeln("Mode: " + config.docChoice.mode);
        
        if (config.docChoice.mode === "new") {
            $.writeln("Creating new document...");
            
            var preset = new DocumentPreset();
            preset.units = RulerUnits.Inches;
            app.preferences.setIntegerPreference("rulerType", 2);
            
            if (config.canvasType === "Large") {
                $.writeln("Creating Large Canvas document with inches");
                preset.width = 300 * 72;
                preset.height = 300 * 72;
                preset.colorMode = DocumentColorSpace.CMYK;
                doc = app.documents.addDocument("Basic CMYK", preset);
                scale = 10;
            } else {
                $.writeln("Creating Standard Canvas document with inches");
                preset.width = 8.5 * 72;
                preset.height = 11 * 72;
                preset.colorMode = DocumentColorSpace.CMYK;
                doc = app.documents.addDocument("Basic CMYK", preset);
                scale = 1;
            }
            $.writeln("Document created: " + doc.name);
            $.writeln("Document ruler units: " + doc.rulerUnits);
        } else {
            $.writeln("Using selected document");
            doc = config.docChoice.document;
            app.activeDocument = doc;
            $.writeln("Document name: " + doc.name);
            try {
                scale = doc.scaleFactor;
                $.writeln("scaleFactor detected: " + scale);
            } catch (e) {
                $.writeln("scaleFactor not available, assuming Standard");
                scale = 1;
            }
        }
        
        $.writeln("Scale factor being used: " + scale);
        $.writeln("Canvas type: " + (scale !== 1 ? "Large Canvas" : "Standard Canvas"));
        $.writeln("Initial artboards count: " + doc.artboards.length);
        
        var hadInitialArtboard = (doc.artboards.length > 0 && config.docChoice.mode === "new");
        $.writeln("Had initial artboard: " + hadInitialArtboard);
        
        // FIXED: Define canvas boundaries based on actual test results
        $.writeln("\n=== CANVAS BOUNDARIES ===");
        
        var canvasBounds;
        if (scale !== 1) {
            // Large Canvas boundaries from your test images
            canvasBounds = {
                minX: -6912,      // From Image 6: left=-6912
                maxX: 9007.2,     // From Image 4: right=9007.2
                minY: -6775.2,    // From Image 6: bottom=-6775.2
                maxY: 9072        // From Image 4: top=9072
            };
        } else {
            // Standard Canvas boundaries from your test images
            canvasBounds = {
                minX: -7632,      // From Image 7/8: left=-7632
                maxX: 8136,       // From Image 9/10: right=8136
                minY: -7416,      // From Image 8/9: bottom=-7416
                maxY: 8352        // From Image 7/10: top=8352
            };
        }
        
        $.writeln("Canvas X range: " + canvasBounds.minX + " to " + canvasBounds.maxX);
        $.writeln("Canvas Y range: " + canvasBounds.minY + " to " + canvasBounds.maxY);
        $.writeln("Canvas width: " + (canvasBounds.maxX - canvasBounds.minX) + " points");
        $.writeln("Canvas height: " + (canvasBounds.maxY - canvasBounds.minY) + " points");
        
        // Add buffer space
        var bufferSpace = (10 * 72) / scale;
        var usableMinX = canvasBounds.minX + bufferSpace;
        var usableMaxX = canvasBounds.maxX - bufferSpace;
        var usableMinY = canvasBounds.minY + bufferSpace;
        var usableMaxY = canvasBounds.maxY - bufferSpace;
        
        $.writeln("Usable X range: " + usableMinX + " to " + usableMaxX);
        $.writeln("Usable Y range: " + usableMinY + " to " + usableMaxY);
        
        // Calculate total space needed and do comprehensive fit check
        var totalSpaceNeeded = calculateSpaceNeeded(config.specs, scale);
        $.writeln("\n=== SPACE REQUIREMENTS ===");
        $.writeln("Total area needed: " + totalSpaceNeeded.totalArea + " sq points");
        $.writeln("Estimated rows needed: " + totalSpaceNeeded.estimatedRows);
        $.writeln("Max single artboard width: " + totalSpaceNeeded.maxRowWidth + " points");
        
        var usableWidth = usableMaxX - usableMinX;
        var usableHeight = usableMaxY - usableMinY;
        
        $.writeln("Usable width: " + usableWidth + " points");
        $.writeln("Usable height: " + usableHeight + " points");
        
        // Check if artboards will fit - need to simulate the layout
        var willFit = true;
        var fitCheckStartX = (scale !== 1) ? -6840 : -6120;
        var fitCheckStartY = (scale !== 1) ? 7704 : 4464;
        var fitCheckCurrentY = fitCheckStartY;
        var spacing = (10 * 72) / scale;
        
        $.writeln("\n=== CHECKING IF LAYOUT WILL FIT ===");
        
        for (var i = 0; i < config.specs.length; i++) {
            var spec = config.specs[i];
            var artboardWidth = (spec.width * 72) / scale;
            var artboardHeight = (spec.height * 72) / scale;
            
            var fitCheckCurrentX = fitCheckStartX;
            var rowStartY = fitCheckCurrentY;
            
            for (var q = 0; q < spec.quantity; q++) {
                var artboardLeft = fitCheckCurrentX;
                var artboardRight = fitCheckCurrentX + artboardWidth;
                var artboardTop = fitCheckCurrentY;
                var artboardBottom = fitCheckCurrentY - artboardHeight;
                
                // Check if we need to wrap
                if (artboardRight > usableMaxX && q > 0) {
                    fitCheckCurrentX = fitCheckStartX;
                    fitCheckCurrentY = rowStartY - artboardHeight - spacing;
                    rowStartY = fitCheckCurrentY;
                    
                    // Recalculate
                    artboardLeft = fitCheckCurrentX;
                    artboardRight = fitCheckCurrentX + artboardWidth;
                    artboardTop = fitCheckCurrentY;
                    artboardBottom = fitCheckCurrentY - artboardHeight;
                }
                
                // Check boundaries
                if (artboardLeft < usableMinX || artboardRight > usableMaxX || 
                    artboardBottom < usableMinY || artboardTop > usableMaxY) {
                    willFit = false;
                    $.writeln("Layout will NOT fit - artboard " + (q+1) + " of material " + spec.material + " exceeds bounds");
                    break;
                }
                
                fitCheckCurrentX += artboardWidth + spacing;
            }
            
            if (!willFit) break;
            
            fitCheckCurrentY = rowStartY - artboardHeight - spacing;
        }
        
        if (!willFit) {
            $.writeln("ERROR: Layout will not fit in available space!");
            
            if (scale === 1 && config.docChoice.mode === "new") {
                // Offer to switch to Large Canvas
                var switchToLarge = confirm(
                    "The artboards will not fit in Standard Canvas.\n\n" +
                    "Switch to Large Canvas mode?\n\n" +
                    "(Large Canvas provides 10x more space)"
                );
                
                if (switchToLarge) {
                    $.writeln("User chose to switch to Large Canvas, restarting...");
                    config.canvasType = "Large";
                    doc.close(SaveOptions.DONOTSAVECHANGES);
                    createArtboards(config);
                    return;
                } else {
                    $.writeln("User cancelled");
                    alert("Operation cancelled. Artboards do not fit in available space.");
                    if (doc && config.docChoice.mode === "new") {
                        doc.close(SaveOptions.DONOTSAVECHANGES);
                    }
                    return;
                }
            } else if (scale !== 1) {
                alert(
                    "Cannot fit artboards in available space.\n\n" +
                    "Even in Large Canvas mode, the layout exceeds maximum dimensions.\n\n" +
                    "Please reduce the number of artboards or use smaller sizes."
                );
                if (doc && config.docChoice.mode === "new") {
                    doc.close(SaveOptions.DONOTSAVECHANGES);
                }
                return;
            } else {
                // Add mode - just show error
                alert("Cannot fit artboards in available space.\n\nPlease reduce the number of artboards or use smaller sizes.");
                return;
            }
        }
        
        $.writeln("Layout check passed - all artboards will fit!");
        
        // FIXED: Use confirmed starting positions from your test images
        var startX, startY;
        
        if (scale !== 1) {
            // Large Canvas: confirmed from Image 2
            startX = -6840;
            startY = 7704;
            $.writeln("Large Canvas starting position (confirmed)");
        } else {
            // Standard Canvas: confirmed from Image 1
            startX = -6120;
            startY = 4464;
            $.writeln("Standard Canvas starting position (confirmed)");
        }
        
        $.writeln("Starting position:");
        $.writeln("  startX: " + startX + " points");
        $.writeln("  startY: " + startY + " points");
        
        var currentY = startY;
        var spacing = (10 * 72) / scale;
        $.writeln("Spacing: " + spacing + " points");
        
        var regLayer = null;
        if (config.regDots !== "None") {
            $.writeln("Creating/getting REG layer...");
            regLayer = getOrCreateLayer(doc, "REG");
            $.writeln("REG layer ready");
        }
        
        if (config.docChoice.mode === "new") {
            $.writeln("Looking for default layer to rename to PRINT...");
            for (var i = 0; i < doc.layers.length; i++) {
                if (doc.layers[i].name === "Layer 1" || doc.layers[i].name === "Unnamed") {
                    doc.layers[i].name = "PRINT";
                    $.writeln("Renamed layer to PRINT");
                    break;
                }
            }
            
            if (!layerExists(doc, "PRINT")) {
                $.writeln("No default layer found, creating PRINT layer...");
                var printLayer = doc.layers.add();
                printLayer.name = "PRINT";
                $.writeln("PRINT layer created");
            }
        }
        
        $.writeln("Processing " + config.specs.length + " material specs...");
        
        // In add mode, analyze existing artboard positions
        var materialPositions = {};
        var lowestY = currentY;
        
        if (config.docChoice.mode === "add" && doc.artboards.length > 0) {
            $.writeln("Add mode: Analyzing existing artboard positions...");
            
            // Build a map of materials and their rightmost positions
            for (var i = 0; i < doc.artboards.length; i++) {
                var ab = doc.artboards[i];
                var abName = ab.name;
                var bounds = ab.artboardRect;
                
                // Extract material name from artboard name
                var match = abName.match(/^(.+?)(?:_Pt(\d+))?$/);
                if (match) {
                    var matName = match[1];
                    var rightEdge = bounds[2];
                    var bottomEdge = bounds[3];
                    
                    if (!materialPositions[matName]) {
                        materialPositions[matName] = {
                            rightmost: rightEdge,
                            bottom: bottomEdge,
                            top: bounds[1],
                            height: bounds[1] - bounds[3]
                        };
                    } else {
                        if (rightEdge > materialPositions[matName].rightmost) {
                            materialPositions[matName].rightmost = rightEdge;
                        }
                    }
                    
                    // Track lowest Y position
                    if (bottomEdge < lowestY) {
                        lowestY = bottomEdge;
                    }
                }
            }
            
            $.writeln("Lowest Y found: " + lowestY);
        }
        
        for (var i = 0; i < config.specs.length; i++) {
            var spec = config.specs[i];
            $.writeln("\n--- Material " + (i+1) + ": " + spec.material + " ---");
            $.writeln("Size: " + spec.width + '"x' + spec.height + '"');
            $.writeln("Quantity: " + spec.quantity);
            $.writeln("Starting part number: " + (spec.startingPartNumber || 1));
            
            var currentX;
            
            // Determine starting position
            if (config.docChoice.mode === "add" && materialPositions[spec.material]) {
                // Adding to existing material - continue to the right
                currentX = materialPositions[spec.material].rightmost + spacing;
                currentY = materialPositions[spec.material].top;
                $.writeln("Continuing existing material at X=" + currentX + ", Y=" + currentY);
            } else if (config.docChoice.mode === "add") {
                // New material in add mode - place below existing
                currentX = startX;
                currentY = lowestY - spacing;
                $.writeln("New material in add mode at X=" + currentX + ", Y=" + currentY);
            } else {
                // New document mode - use standard positioning
                currentX = startX;
                $.writeln("New document mode at X=" + currentX + ", Y=" + currentY);
            }
            
            var artboardWidth = (spec.width * 72) / scale;
            var artboardHeight = (spec.height * 72) / scale;
            
            $.writeln("Artboard dimensions (points): " + artboardWidth + "x" + artboardHeight);
            
            var rowStartY = currentY;
            var maxHeightInRow = artboardHeight;
            
            for (var q = 0; q < spec.quantity; q++) {
                $.writeln("Creating artboard " + (q+1) + " of " + spec.quantity);
                
                // Calculate where this artboard would be placed
                var artboardLeft = currentX;
                var artboardRight = currentX + artboardWidth;
                var artboardTop = currentY;
                var artboardBottom = currentY - artboardHeight;
                
                $.writeln("Proposed position:");
                $.writeln("  Left: " + artboardLeft + " (min: " + usableMinX + ")");
                $.writeln("  Right: " + artboardRight + " (max: " + usableMaxX + ")");
                $.writeln("  Top: " + artboardTop + " (max: " + usableMaxY + ")");
                $.writeln("  Bottom: " + artboardBottom + " (min: " + usableMinY + ")");
                
                // Check if we need to wrap to next row
                if (artboardRight > usableMaxX && q > 0) {
                    $.writeln("Would exceed right boundary, wrapping to next row");
                    currentX = startX;
                    currentY = rowStartY - maxHeightInRow - spacing;
                    rowStartY = currentY;
                    
                    // Recalculate positions
                    artboardLeft = currentX;
                    artboardRight = currentX + artboardWidth;
                    artboardTop = currentY;
                    artboardBottom = currentY - artboardHeight;
                    
                    $.writeln("Wrapped position:");
                    $.writeln("  Left: " + artboardLeft);
                    $.writeln("  Right: " + artboardRight);
                    $.writeln("  Top: " + artboardTop);
                    $.writeln("  Bottom: " + artboardBottom);
                }
                
                // Final boundary check
                if (artboardLeft < usableMinX || artboardRight > usableMaxX || 
                    artboardBottom < usableMinY || artboardTop > usableMaxY) {
                    
                    var errorMsg = "Cannot create artboard - exceeds canvas boundaries:\n";
                    if (artboardLeft < usableMinX) errorMsg += "  Left edge too far left\n";
                    if (artboardRight > usableMaxX) errorMsg += "  Right edge too far right\n";
                    if (artboardBottom < usableMinY) errorMsg += "  Bottom edge too low\n";
                    if (artboardTop > usableMaxY) errorMsg += "  Top edge too high\n";
                    
                    $.writeln("ERROR: " + errorMsg);
                    throw new Error(errorMsg);
                }
                
                var artboardName;
                var partNumber = spec.startingPartNumber ? spec.startingPartNumber + q : q + 1;
                
                // Determine if this is a single artboard or part of a series
                var totalParts = spec.startingPartNumber ? spec.startingPartNumber + spec.quantity - 1 : spec.quantity;
                
                if (totalParts === 1) {
                    artboardName = spec.material;
                } else {
                    artboardName = spec.material + "_Pt" + partNumber;
                }
                
                $.writeln("Creating artboard: " + artboardName);
                $.writeln("Final position: [" + artboardLeft + ", " + artboardTop + ", " + artboardRight + ", " + artboardBottom + "]");
                
                var ab = doc.artboards.add([artboardLeft, artboardTop, artboardRight, artboardBottom]);
                ab.name = artboardName;
                
                $.writeln("Artboard created successfully!");
                $.writeln("Actual bounds: " + ab.artboardRect);
                
                if (config.regDots !== "None") {
                    $.writeln("Creating reg dots...");
                    createRegDots(doc, regLayer, ab, config.regDots, scale);
                    $.writeln("Reg dots created");
                } else {
                    $.writeln("Skipping reg dots (None selected)");
                }
                
                currentX += artboardWidth + spacing;
                $.writeln("Next X position: " + currentX);
            }
            
            // Move to next vertical position for next material (after all rows of this material)
            currentY = rowStartY - maxHeightInRow - spacing;
            $.writeln("Next material Y position: " + currentY);
        }
        
        if (hadInitialArtboard && doc.artboards.length > 1) {
            $.writeln("\nRemoving initial placeholder artboard at index 0...");
            $.writeln("Artboards before removal: " + doc.artboards.length);
            doc.artboards.remove(0);
            $.writeln("Artboards after removal: " + doc.artboards.length);
        }
        
        try {
            app.executeMenuCommand("fitall");
            $.writeln("Fit complete");
        } catch (e) {
            $.writeln("Fit command failed: " + e.toString());
        }
        
        if (config.docChoice.mode === "new") {
            $.writeln("\n=== SAVING DOCUMENT ===");
            saveAsPrimePDF(doc, config.docChoice.primeFolder);
        }
        
        $.writeln("\n=== ARTBOARD CREATION COMPLETE ===");
        
    } catch (e) {
        $.writeln("!!! ERROR !!!");
        $.writeln("Error message: " + e.toString());
        $.writeln("Error line: " + e.line);
        $.writeln("Stack: " + (e.stack ? e.stack : "undefined"));
        alert("Error: " + e.toString() + "\n\nLine: " + e.line + "\n\nCheck ESTK console for details");
    }
}

function saveAsPrimePDF(doc, primeFolder) {
    try {
        var folder = Folder(primeFolder);
        if (!folder.exists) {
            throw new Error("PRIME folder does not exist: " + primeFolder);
        }
        
        var proofFile = findProofFile(primeFolder);
        if (!proofFile) {
            $.writeln("No _Proof.pdf file found, using default name");
            var saveFile = new File(primeFolder + "/PRIME.pdf");
        } else {
            var originalName = decodeURI(proofFile.name);
            var nameWithoutExtension = originalName.replace(/\.[^\.]+$/, '');
            var baseName = nameWithoutExtension.replace(/\s*_Proof$/i, "");
            var newName = baseName + "_PRIME";
            
            $.writeln("Proof file found: " + originalName);
            $.writeln("Base name: " + baseName);
            $.writeln("New name: " + newName);
            
            var saveFile = new File(primeFolder + "/" + newName + ".pdf");
        }
        
        var pdfOptions = new PDFSaveOptions();
        pdfOptions.pdfPreset = "[Illustrator Default]";
        pdfOptions.acrobatLayers = true;
        pdfOptions.colorCompression = CompressionQuality.None;
        pdfOptions.compatibility = PDFCompatibility.ACROBAT7;
        
        doc.saveAs(saveFile, pdfOptions);
        $.writeln("Document saved: " + saveFile.fsName);
        
    } catch (e) {
        $.writeln("Save error: " + e.toString());
        throw e;
    }
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

function getOrCreateLayer(doc, layerName) {
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === layerName) {
            return doc.layers[i];
        }
    }
    var newLayer = doc.layers.add();
    newLayer.name = layerName;
    return newLayer;
}

function layerExists(doc, layerName) {
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === layerName) {
            return true;
        }
    }
    return false;
}

function createRegDots(doc, layer, artboard, placement, scale) {
    var bounds = artboard.artboardRect;
    var left = bounds[0];
    var top = bounds[1];
    var right = bounds[2];
    var bottom = bounds[3];
    
    var width = right - left;
    var height = top - bottom;
    
    var dotRadius = (0.5 * 72) / scale;
    
    var dotPositions = [];
    
    if (placement === "TopBottom") {
        var effectiveLeft = left + dotRadius;
        var effectiveRight = right - dotRadius;
        var effectiveWidth = effectiveRight - effectiveLeft;
        var spacing = effectiveWidth / 4;
        
        for (var i = 0; i < 5; i++) {
            var x = effectiveLeft + (i * spacing);
            var y = top - dotRadius;
            dotPositions.push({x: x, y: y});
        }
        
        dotPositions[1].y = bottom + dotRadius;
        dotPositions[3].y = bottom + dotRadius;
        
    } else {
        var effectiveTop = top - dotRadius;
        var effectiveBottom = bottom + dotRadius;
        var effectiveHeight = effectiveTop - effectiveBottom;
        var spacing = effectiveHeight / 4;
        
        for (var i = 0; i < 5; i++) {
            var x = left + dotRadius;
            var y = effectiveTop - (i * spacing);
            dotPositions.push({x: x, y: y});
        }
        
        dotPositions[1].x = right - dotRadius;
        dotPositions[3].x = right - dotRadius;
    }
    
    for (var i = 0; i < dotPositions.length; i++) {
        createSingleRegDot(doc, layer, dotPositions[i].x, dotPositions[i].y, scale);
    }
}

function createSingleRegDot(doc, layer, centerX, centerY, scale) {
    var outerRadius = (0.5 * 72) / scale;
    var outerCircle = layer.pathItems.ellipse(
        centerY + outerRadius,
        centerX - outerRadius,
        outerRadius * 2,
        outerRadius * 2
    );
    outerCircle.filled = false;
    outerCircle.stroked = false;
    
    var innerRadius = (0.125 * 72) / scale;
    var innerCircle = layer.pathItems.ellipse(
        centerY + innerRadius,
        centerX - innerRadius,
        innerRadius * 2,
        innerRadius * 2
    );
    
    var cmykBlack = new CMYKColor();
    cmykBlack.cyan = 0;
    cmykBlack.magenta = 0;
    cmykBlack.yellow = 0;
    cmykBlack.black = 100;
    
    innerCircle.filled = true;
    innerCircle.fillColor = cmykBlack;
    innerCircle.stroked = false;
    
    var clipGroup = layer.groupItems.add();
    innerCircle.moveToBeginning(clipGroup);
    outerCircle.moveToBeginning(clipGroup);
    clipGroup.clipped = true;
}