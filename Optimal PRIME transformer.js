/*
@METADATA
{
  "name": "Optimal PRIME Transformer",
  "description": "Generate Production Files From a PRIME",
  "version": "2.1",
  "target": "illustrator",
  "tags": ["Optimal", "Prime", "processors", "scaleFactor"]
}
@END_METADATA
*/

// ==== PROD NAME PARSER v1.0 ====
// Shared production file/artboard name parser. Keep this block IDENTICAL in:
//   File Check.js, OpenProdFilesForReview.js, Link Namer.js,
//   Sort Linked Files By Pt Number.js, Optimal PRIME transformer.js
// Standard:  {base}_{TYPE}_{Material}[_BlkBk][_FRONT|_BACK]_{Parts}[_x#].pdf
//   TYPE  = PRINT | CUT | PnC | CutOnly
//   Parts = Pt1 | Pt3-5 | Pt3-5and8 (CUT files may list several parts)
//   x#    = number of times the sheet is run (always last)
// Legacy names still parse: 3mmACMBLK, BlkBk_6mmACM, _Part2.
var ProdName = (function () {
    var TYPE_TOKENS = [
        ['_CUTONLY_', 'CutOnly'],
        ['_PRINT_', 'PRINT'],
        ['_PNC_', 'PnC'],
        ['_CUT_', 'CUT']
    ];

    function trim(s) {
        return String(s).replace(/^\s+|\s+$/g, '');
    }

    function stripExt(name) {
        var n = String(name).replace(/%20/g, ' ');
        return n.replace(/\.(ai|pdf|eps|svg)$/i, '');
    }

    // "3-5and8" -> [3,4,5,8]; returns null when the text is not a valid part list
    function expandParts(spec) {
        var out = [];
        var chunks = String(spec).split(/and/i);
        for (var i = 0; i < chunks.length; i++) {
            var m = trim(chunks[i]).match(/^(\d+)(?:\s*-\s*(\d+))?$/);
            if (!m) { return null; }
            var a = parseInt(m[1], 10);
            var b = m[2] ? parseInt(m[2], 10) : a;
            if (b < a) { var t = a; a = b; b = t; }
            if (b - a > 999) { return null; }
            for (var n = a; n <= b; n++) { out.push(n); }
        }
        return out;
    }

    // [3,4,5,8,10,11] -> "3-5and8and10-11"
    function formatParts(parts) {
        var list = [];
        var seen = {};
        for (var i = 0; i < parts.length; i++) {
            var v = parseInt(parts[i], 10);
            if (!isNaN(v) && !seen[v]) { seen[v] = true; list.push(v); }
        }
        list.sort(function (x, y) { return x - y; });
        var out = [];
        var i2 = 0;
        while (i2 < list.length) {
            var start = list[i2];
            var end = start;
            while (i2 + 1 < list.length && list[i2 + 1] === end + 1) { i2++; end = list[i2]; }
            out.push(start === end ? String(start) : (start + '-' + end));
            i2++;
        }
        return out.join('and');
    }

    // Parse everything after TYPE (or a whole artboard / material name)
    function parseTail(tail) {
        var r = {
            material: '', blkbk: false, legacyBlk: false, side: '',
            parts: [], partsText: '', mult: 1, hasMult: false
        };
        var raw = String(tail).split('_');
        var tk = [];
        for (var i = 0; i < raw.length; i++) {
            var s = trim(raw[i]);
            if (s !== '') { tk.push(s); }
        }
        // Multiple: last token xN
        if (tk.length > 1) {
            var mm = tk[tk.length - 1].match(/^x(\d+)$/i);
            if (mm) {
                r.mult = Math.max(1, parseInt(mm[1], 10));
                r.hasMult = true;
                tk.pop();
            }
        }
        // Parts: PtN / PtN-M / PtNandM (legacy PartN)
        if (tk.length > 1) {
            var pm = tk[tk.length - 1].match(/^(?:Pt|Part)\s*(\d[\d\s\-]*(?:and\s*\d[\d\s\-]*)*)$/i);
            if (pm) {
                var parts = expandParts(pm[1]);
                if (parts) {
                    r.parts = parts;
                    r.partsText = formatParts(parts);
                    tk.pop();
                }
            }
        }
        // Side and BlkBk tokens (any position after the first, BlkBk anywhere)
        var mat = [];
        for (var j = 0; j < tk.length; j++) {
            var u = tk[j].toUpperCase();
            if (u === 'BLKBK' || u === 'BLK') {
                r.blkbk = true;
                if (j === 0 || u === 'BLK') { r.legacyBlk = true; }
            } else if ((u === 'FRONT' || u === 'BACK') && j > 0) {
                r.side = u;
            } else {
                mat.push(tk[j]);
            }
        }
        var material = trim(mat.join('_'));
        // Legacy glued suffix: 3mmACMBLK / 3mmACM BlkBk / 3mmACM-BLK
        var lg = material.match(/^(.*\S)[\s\-]*(?:BLKBK|BLK)$/i);
        if (lg && !r.blkbk) {
            material = trim(lg[1]);
            r.blkbk = true;
            r.legacyBlk = true;
        }
        r.material = material;
        return r;
    }

    function copyInto(dst, src) {
        for (var k in src) { if (src.hasOwnProperty(k)) { dst[k] = src[k]; } }
        return dst;
    }

    // Full production file name
    function parseFileName(fileName) {
        var name = stripExt(fileName);
        var upper = name.toUpperCase();
        var r = { name: String(fileName), base: name, type: '', isPrime: false };
        for (var t = 0; t < TYPE_TOKENS.length; t++) {
            var idx = upper.indexOf(TYPE_TOKENS[t][0]);
            if (idx !== -1) {
                r.type = TYPE_TOKENS[t][1];
                r.base = name.substring(0, idx);
                copyInto(r, parseTail(name.substring(idx + TYPE_TOKENS[t][0].length)));
                return r;
            }
        }
        r.isPrime = /PRIME/i.test(name);
        copyInto(r, parseTail(''));
        return r;
    }

    // Artboard name (PRIME) or material name from packing data
    function parseArtboardName(abName) {
        return parseTail(abName);
    }

    // Normalized material key for comparisons ("3mm ACM" == "3mmACM" == "3MMACM")
    function matKey(material) {
        return String(material).toUpperCase().replace(/[\s_\-]+/g, '');
    }

    function baseKey(base) {
        return String(base).toUpperCase().replace(/[\s_\-]+/g, '');
    }

    // Key used to pair a PRINT part with its CUT part. FRONT, BACK and no side are distinct.
    function pairKey(p, part) {
        var side = (p.side === 'BACK' || p.side === 'FRONT') ? p.side : '';
        var pt = (part === undefined || part === null) ? (p.parts.length ? p.parts[0] : 0) : part;
        return baseKey(p.base) + '|' + matKey(p.material) + '|' + (p.blkbk ? 'B' : '') + '|' + side + '|' + pt;
    }

    // Older files: a _FRONT print whose CUT has no side tag. Only used as a
    // fallback when no _FRONT cut exists (a one-sided Pt1 must not grab it).
    function pairKeyNoSide(p, part) {
        var pt = (part === undefined || part === null) ? (p.parts.length ? p.parts[0] : 0) : part;
        return baseKey(p.base) + '|' + matKey(p.material) + '|' + (p.blkbk ? 'B' : '') + '||' + pt;
    }

    // Readable short label, e.g. "PRINT 3mmACM BlkBk Pt8 x2"
    function label(p) {
        var s = [];
        if (p.type) { s.push(p.type); }
        if (p.material) { s.push(p.material); }
        if (p.blkbk) { s.push('BlkBk'); }
        if (p.side) { s.push(p.side); }
        if (p.parts.length) { s.push('Pt' + p.partsText); }
        if (p.hasMult) { s.push('x' + p.mult); }
        return s.join(' ');
    }

    // Rebuild a standard artboard/file tail from parts
    function buildTail(material, blkbk, side, parts, mult) {
        var s = [material];
        if (blkbk) { s.push('BlkBk'); }
        if (side) { s.push(side); }
        if (parts && parts.length) { s.push('Pt' + formatParts(parts)); }
        if (mult && mult > 1) { s.push('x' + mult); }
        return s.join('_');
    }

    return {
        parseFileName: parseFileName,
        parseArtboardName: parseArtboardName,
        expandParts: expandParts,
        formatParts: formatParts,
        matKey: matKey,
        baseKey: baseKey,
        pairKey: pairKey,
        pairKeyNoSide: pairKeyNoSide,
        label: label,
        buildTail: buildTail
    };
})();
// ==== END PROD NAME PARSER ====

(function() {
    // Global variables
    var doc;
    var originalDocPath; // Store original document path
    var parentSavePath; // Store the parent path for all saves
    var overwriteAll = null; // null, true, or false
    var skipAll = false;
    var scaleFactor = 1; // Add scale factor tracking
    var useFolderOrganization = false; // Track if we should organize into folders
    var materialData = {}; // Store material information for folder organization
    var sharedCuts = { skip: {}, tail: {} }; // Shared CUT files (see detectSharedCuts)
    
    // Main execution
    try {
        // Check if document is open
        if (app.documents.length === 0) {
            alert("Please open a document before running this script.");
            return;
        }

        // Run Cut Path Separator first (kept as its own file - single source of truth).
        // $.evalFile sets $.fileName to this script's path, so CPS is found beside it
        // in whatever folder the launcher downloaded to. Edits to CPS flow through automatically.
        var cpsFile = new File(File($.fileName).parent + "/Cut Path Separator.js");
        if (cpsFile.exists) {
            $.evalFile(cpsFile);
        } else {
            alert("Cut Path Separator.js not found next to this script - skipping pre-step.");
        }

        doc = app.activeDocument;
        
        // Get the scale factor for large canvas handling
        scaleFactor = doc.scaleFactor || 1; // Default to 1 if not defined
        
        // Store the original document path and calculate parent save path once
        originalDocPath = doc.fullName;
        parentSavePath = originalDocPath.parent.parent;
        
        // Check if Cut Path Separator has been run
        var cpsStatus = checkForCPSProcessing();
        
        if (!cpsStatus.hasProcessedLayers) {
            var message = "It appears the Cut Path Separator script has not been run on this document.\n\n";
            
            if (cpsStatus.foundLayers.length > 0) {
                message += "Found CPS layer(s): " + cpsStatus.foundLayers.join(", ") + "\n";
                message += "But these layers appear to be empty.\n\n";
            } else {
                message += "No CPS-created layers found.\n\n";
            }
            
            message += "It is recommended to run the Cut Path Separator script first to:\n";
            message += "- Organize cut/print data into proper layers\n";
            message += "- Standardize stroke weights\n";
            message += "- Check for duplicate paths\n";
            message += "- Enable overprint settings\n\n";
            message += "Do you want to continue anyway?";
            
            var continueProcessing = confirm(message);
            if (!continueProcessing) {
                return; // Exit the script
            }
        }
        
        // Check for PRIME in filename
        if (!checkForPrime()) {
            return;
        }
        
        // Check artboard names
        if (!checkArtboardNames()) {
            return;
        }
        
        // Check for illegal characters in artboard names
        if (!checkArtboardCharacters()) {
            return;
        }
        
		// Check if we should use folder organization
        if (doc.artboards.length > 1) {
            // Validate material naming and get summary in one step
            var materialValidation = validateMaterialNaming();
            
            if (!materialValidation.isValid) {
                var errorMessage = "*** ARTBOARD NAMING ISSUES DETECTED:\n\n";
                errorMessage += materialValidation.errors.join("\n") + "\n\n";
                errorMessage += "Correct format: 'Material[_BlkBk][_FRONT|_BACK]_PtX[_xN]' where:\n";
                errorMessage += "- Material can contain letters, numbers, spaces, and dashes\n";
                errorMessage += "- _PtX is optional for single parts (X can be any number)\n";
                errorMessage += "- _xN (optional, last) = sheet is run N times, e.g. 3mmACM_Pt7_x3\n";
                errorMessage += "- _BlkBk = black-back stock, e.g. 3mmACM_BlkBk_Pt1\n";
                errorMessage += "- Part numbers must be unique for each material\n\n";
                errorMessage += "Please fix the artboard names before running the script again.";
                
                alert(errorMessage);
                return;
            } else {
                // Show folder organization dialog
                var folderChoice = showFolderOrganizationDialog(materialValidation.materialData);
                
                if (folderChoice === null) {
                    return; // User cancelled
                } else if (folderChoice === true) {
                    useFolderOrganization = true;
                    materialData = materialValidation.materialData;
                } else {
                    useFolderOrganization = false;
                    materialData = {};
                }
            }
        }
		
        // Show artboard selection dialog so user can re-export specific artboards
        var selectedArtboards = showArtboardSelectionDialog();
        if (selectedArtboards === null) {
            return; // User cancelled
        }

        // Offer to combine identical cut layouts into shared CUT files (one dialog,
        // only shown when matches are found)
        sharedCuts = detectSharedCuts(selectedArtboards);

        // Persist Cut Path Separator's layer changes to the original file BEFORE processing.
        // The artboard loop reopens originalDocPath from disk between artboards, so the on-disk
        // file must already contain CPS's separated layers or those changes are lost from the
        // original. Use a format-aware saveAs (bare doc.save() throws on non-.ai working files).
        try {
            var origExt = (originalDocPath.name.match(/\.([^.]+)$/) || ["", ""])[1].toLowerCase();
            if (origExt === "pdf") {
                var pdfSaveOpts = new PDFSaveOptions();
                pdfSaveOpts.preserveEditability = true;
                doc.saveAs(originalDocPath, pdfSaveOpts);
            } else if (origExt === "eps") {
                doc.saveAs(originalDocPath, new EPSSaveOptions());
            } else {
                // .ai (and anything else) - mirror the options the loop already uses successfully
                var aiSaveOpts = new IllustratorSaveOptions();
                aiSaveOpts.compatibility = Compatibility.ILLUSTRATOR24;
                aiSaveOpts.pdfCompatible = true;
                doc.saveAs(originalDocPath, aiSaveOpts);
            }
        } catch (eSave) {
            alert("Could not persist Cut Path Separator's layer changes to the original file.\n" +
                  "The exported files are still correct, but the original PRIME will not retain the separation.\n\n" +
                  "Error: " + eSave.message);
        }

        // Process only the selected artboards
        processAllArtboards(selectedArtboards);
        
        alert("Processing complete!" + (scaleFactor !== 1 ? "\nLarge Canvas scaling factor (" + scaleFactor + ") has been applied." : ""));
        
    } catch (e) {
        alert("Error occurred: " + e.toString() + "\nLine: " + e.line);
    }
	
	function showFolderOrganizationDialog(materialValidation) {
        var dialog = new Window("dialog", "Folder Organization");
        dialog.orientation = "column";
        dialog.alignChildren = "fill";
        dialog.spacing = 15;
        dialog.margins = 20;
        dialog.preferredSize.width = 400;
        
        // Header
        var headerText = dialog.add("statictext", undefined, "This document has " + doc.artboards.length + " artboards.");
        headerText.graphics.font = ScriptUI.newFont("dialog", "bold", 12);
        
        // Materials summary
        var summaryGroup = dialog.add("group");
        summaryGroup.orientation = "column";
        summaryGroup.alignChildren = "fill";
        
        summaryGroup.add("statictext", undefined, "Materials detected:");
        
        var materialsList = summaryGroup.add("edittext", undefined, "", {multiline: true, readonly: true});
        materialsList.preferredSize.height = 120;
        
        var materialsText = "";
        for (var materialName in materialValidation) {
            var count = materialValidation[materialName].artboards.length;
            materialsText += "- " + materialName + ": " + count + " part" + (count > 1 ? "s" : "") + "\n";
            if (count > 10 && !materialUsesCutContour(materialName)) {
                materialsText += "  > Will create PRINT and CUT subfolders\n";
            }
        }
        materialsList.text = materialsText;
        
        // Instructions
        var instructText = dialog.add("statictext", undefined, "Choose how to organize the output files:");
        instructText.graphics.font = ScriptUI.newFont("dialog", "bold", 11);
        
        // Button group
        var buttonGroup = dialog.add("group");
        buttonGroup.alignment = "center";
        buttonGroup.spacing = 10;
        
        var foldersButton = buttonGroup.add("button", undefined, "Create Folders");
        foldersButton.preferredSize.width = 100;
        
        var rootButton = buttonGroup.add("button", undefined, "Root Directory");
        rootButton.preferredSize.width = 100;
        
        var cancelButton = buttonGroup.add("button", undefined, "Cancel");
        cancelButton.preferredSize.width = 80;
        
        // Button handlers
        var result = null;
        
        foldersButton.onClick = function() {
            result = true;
            dialog.close();
        };
        
        rootButton.onClick = function() {
            result = false;
            dialog.close();
        };
        
        cancelButton.onClick = function() {
            result = null;
            dialog.close();
        };
        
        // Show dialog
        dialog.show();
        
        return result;
    }
    
    // Show a dialog letting the user pick which artboards to export.
    // Returns an array of selected artboard indices, or null if cancelled.
    function showArtboardSelectionDialog() {
        var artboardCount = doc.artboards.length;

        // Single artboard - no need to ask, just export it
        if (artboardCount === 1) {
            return [0];
        }

        var dialog = new Window("dialog", "Select Artboards to Export");
        dialog.orientation = "column";
        dialog.alignChildren = "fill";
        dialog.spacing = 12;
        dialog.margins = 20;
        dialog.preferredSize.width = 400;

        // Header
        var headerText = dialog.add("statictext", undefined, "Choose which artboards to export:");
        headerText.graphics.font = ScriptUI.newFont("dialog", "bold", 12);

        // Multi-select listbox - auto-scrolls when there are many artboards
        // Height grows with the number of artboards up to a cap; beyond that, it scrolls.
        var listHeight = Math.min(artboardCount * 20 + 10, 500);
        var list = dialog.add("listbox", undefined, [], {multiselect: true});
        list.preferredSize.width = 360;
        list.preferredSize.height = listHeight;

        for (var i = 0; i < artboardCount; i++) {
            list.add("item", (i + 1) + ":  " + doc.artboards[i].name);
        }

        // Default: all selected
        var initialSelection = [];
        for (var i = 0; i < artboardCount; i++) initialSelection.push(i);
        list.selection = initialSelection;

        // Help text - clarifies the multi-select interaction
        var helpText = dialog.add("statictext", undefined,
            "Click to select. Ctrl/Cmd+Click toggles individual items. Shift+Click selects a range.");
        helpText.graphics.font = ScriptUI.newFont("dialog", "italic", 10);

        // Selection count display
        var countText = dialog.add("statictext", undefined, "");
        countText.graphics.font = ScriptUI.newFont("dialog", "bold", 10);

        function updateCount() {
            var n = 0;
            if (list.selection) {
                // selection can be a single item or an array, depending on multiselect state
                n = list.selection.length !== undefined ? list.selection.length : 1;
            }
            countText.text = n + " of " + artboardCount + " artboards selected";
        }
        updateCount();
        list.onChange = updateCount;

        // Select All / Deselect All row
        var toggleGroup = dialog.add("group");
        toggleGroup.alignment = "center";
        toggleGroup.spacing = 10;

        var selectAllBtn = toggleGroup.add("button", undefined, "Select All");
        selectAllBtn.preferredSize.width = 100;
        var deselectAllBtn = toggleGroup.add("button", undefined, "Deselect All");
        deselectAllBtn.preferredSize.width = 100;

        selectAllBtn.onClick = function() {
            var all = [];
            for (var i = 0; i < artboardCount; i++) all.push(i);
            list.selection = all;
            updateCount();
        };
        deselectAllBtn.onClick = function() {
            list.selection = null;
            updateCount();
        };

        // OK / Cancel row
        var buttonGroup = dialog.add("group");
        buttonGroup.alignment = "center";
        buttonGroup.spacing = 10;

        var okBtn = buttonGroup.add("button", undefined, "Export Selected");
        okBtn.preferredSize.width = 130;
        var cancelBtn = buttonGroup.add("button", undefined, "Cancel");
        cancelBtn.preferredSize.width = 80;

        var result = null;

        okBtn.onClick = function() {
            var selected = [];
            if (list.selection) {
                if (list.selection.length !== undefined) {
                    // Array of items
                    for (var i = 0; i < list.selection.length; i++) {
                        selected.push(list.selection[i].index);
                    }
                } else {
                    // Single item
                    selected.push(list.selection.index);
                }
            }
            if (selected.length === 0) {
                alert("Please select at least one artboard to export.");
                return;
            }
            // Sort indices ascending to preserve artboard order
            selected.sort(function(a, b) { return a - b; });
            result = selected;
            dialog.close();
        };

        cancelBtn.onClick = function() {
            result = null;
            dialog.close();
        };

        dialog.show();
        return result;
    }

    // Check if Cut Path Separator has been run by looking for CPS-created layers with content
    function checkForCPSProcessing() {
        var cpsLayerNames = ["CutThrough2-Outside", "CutThrough1-Inside", "PinMountHoles-Inside", "PinHoles", "CutThrough-Knifecut", "CutContour", "Spot1"];
        var foundLayers = [];
        var layersWithContent = [];
        
        // Check each CPS layer
        for (var i = 0; i < doc.layers.length; i++) {
            var layer = doc.layers[i];
            var layerName = layer.name;
            
            // Check if this is a CPS-created layer
            for (var j = 0; j < cpsLayerNames.length; j++) {
                if (layerName === cpsLayerNames[j]) {
                    foundLayers.push(layerName);
                    
                    // Check if layer has content using the existing layerHasContent function
                    if (layerHasContent(layer)) {
                        layersWithContent.push(layerName);
                    }
                    break;
                }
            }
        }
        
        return {
            foundLayers: foundLayers,
            layersWithContent: layersWithContent,
            hasProcessedLayers: layersWithContent.length > 0
        };
    }
    
    // Validate material naming format for folder organization.
    // Artboard names follow the production naming standard (PROD NAME PARSER):
    //   Material[_BlkBk][_FRONT|_BACK][_Pt#][_x#]   e.g. 3mmACM_Pt7_x3, 3mmACM_BlkBk_Pt1
    function validateMaterialNaming() {
        var materialData = {};
        var errors = [];
        var isValid = true;
        var seenParts = {}; // group|side -> { partNumber: true }

        for (var i = 0; i < doc.artboards.length; i++) {
            var artboardName = doc.artboards[i].name;
            if (typeof artboardName === 'string') {
                artboardName = artboardName.replace(/^\s+|\s+$/g, ''); // Manual trim for compatibility
            }

            // Allowed characters: letters, numbers, spaces, underscore, dash
            if (!/^[a-zA-Z0-9\s_\-]+$/.test(artboardName)) {
                var uniqueChars = [];
                for (var charIndex = 0; charIndex < artboardName.length; charIndex++) {
                    var currentChar = artboardName.charAt(charIndex);
                    if (!/[a-zA-Z0-9\s_\-]/.test(currentChar)) {
                        var found = false;
                        for (var u = 0; u < uniqueChars.length; u++) {
                            if (uniqueChars[u] === currentChar) { found = true; break; }
                        }
                        if (!found) { uniqueChars.push(currentChar); }
                    }
                }
                errors.push("Artboard " + (i + 1) + " '" + artboardName + "': Contains illegal characters: " + uniqueChars.join(" "));
                isValid = false;
                continue;
            }

            var parsed = ProdName.parseArtboardName(artboardName);
            if (!parsed.material) {
                errors.push("Artboard " + (i + 1) + " '" + artboardName + "': No material name found");
                isValid = false;
                continue;
            }
            // A Pt/x token the parser could not read ends up inside the material name
            if (/(^|_)(Pt|Part)\s*\d/i.test(parsed.material) || /(^|_)x\d+(_|$)/i.test(parsed.material)) {
                errors.push("Artboard " + (i + 1) + " '" + artboardName + "': Part or multiple not in standard form (e.g. 3mmACM_Pt1, 3mmACM_Pt1_x2)");
                isValid = false;
                continue;
            }
            if (parsed.parts.length > 1) {
                errors.push("Artboard " + (i + 1) + " '" + artboardName + "': An artboard is one part - use a single Pt number");
                isValid = false;
                continue;
            }

            var materialName = getMaterialNameFromArtboardName(artboardName);
            var partNumber = parsed.parts.length ? String(parsed.parts[0]) : null;

            // Initialize material data if first occurrence
            if (!materialData[materialName]) {
                materialData[materialName] = {
                    artboards: [],
                    partNumbers: [],
                    originalCase: materialName // Store first occurrence case
                };
            }

            // Check for duplicate part numbers within the same material (and side)
            if (partNumber !== null) {
                var dupKey = materialName + "|" + parsed.side;
                if (!seenParts[dupKey]) { seenParts[dupKey] = {}; }
                if (seenParts[dupKey][partNumber]) {
                    errors.push("Material '" + materialName + (parsed.side ? " " + parsed.side : "") + "': Duplicate part number " + partNumber);
                    isValid = false;
                    continue;
                }
                seenParts[dupKey][partNumber] = true;
                materialData[materialName].partNumbers.push(partNumber);
            }

            // Add artboard to material data
            materialData[materialName].artboards.push({
                index: i,
                name: artboardName,
                partNumber: partNumber
            });
        }

        return {
            isValid: isValid,
            errors: errors,
            materialData: materialData
        };
    }

    // ---------------------------------------------------------------
    // SHARED CUT DETECTION
    // Artboards of the same material whose cut layers are identical (same
    // artboard size, same paths in the same spots) can share ONE cut file,
    // named with all their parts: CUT_3mmACM_Pt3-5and8. The user confirms
    // which groups to combine in a single dialog.
    // ---------------------------------------------------------------
    function isCutLayerForSharing(layerName) {
        if (layerName === "CutContour" || layerName === "CutContour_PRF") { return false; }
        if (isPinHoleLayerName(layerName)) { return true; }
        return String(layerName).toLowerCase().indexOf("cut") !== -1;
    }

    function collectCutBounds() {
        var out = [];
        function walk(container, layerName) {
            var items;
            try { items = container.pageItems; } catch (e) { return; }
            for (var i = 0; i < items.length; i++) {
                var it = items[i];
                try {
                    if (it.typename === "GroupItem") {
                        walk(it, layerName);
                    } else if (it.typename === "PathItem" || it.typename === "CompoundPathItem") {
                        out.push({ layer: layerName, b: it.geometricBounds });
                    }
                } catch (e2) {}
            }
        }
        function walkLayer(layer, topName) {
            walk(layer, topName);
            try {
                for (var s = 0; s < layer.layers.length; s++) { walkLayer(layer.layers[s], topName); }
            } catch (e3) {}
        }
        for (var l = 0; l < doc.layers.length; l++) {
            if (isCutLayerForSharing(doc.layers[l].name)) {
                walkLayer(doc.layers[l], doc.layers[l].name);
            }
        }
        return out;
    }

    function detectSharedCuts(selectedIndices) {
        var result = { skip: {}, tail: {} };
        if (!selectedIndices || selectedIndices.length < 2) { return result; }

        var cuts = collectCutBounds();
        if (cuts.length === 0) { return result; }
        var TOL = 0.5; // points in the document (x scale factor in real size)

        // Per-artboard cut layout relative to the artboard's top-left corner
        var boards = [];
        for (var s = 0; s < selectedIndices.length; s++) {
            var idx = selectedIndices[s];
            var ab = doc.artboards[idx];
            var r = ab.artboardRect;
            var parsed = ProdName.parseArtboardName(ab.name);
            if (!parsed.material || parsed.parts.length !== 1) { continue; }
            var entries = [];
            for (var c = 0; c < cuts.length; c++) {
                var b = cuts[c].b;
                var cx = (b[0] + b[2]) / 2;
                var cy = (b[1] + b[3]) / 2;
                if (cx >= r[0] && cx <= r[2] && cy <= r[1] && cy >= r[3]) {
                    entries.push({ layer: cuts[c].layer, l: b[0] - r[0], t: r[1] - b[1], w: b[2] - b[0], h: b[1] - b[3] });
                }
            }
            if (entries.length === 0) { continue; }
            entries.sort(function (p, q) {
                var dt = Math.round(p.t) - Math.round(q.t);
                if (dt !== 0) { return dt; }
                return p.l - q.l;
            });
            boards.push({
                index: idx,
                parsed: parsed,
                w: r[2] - r[0],
                h: r[1] - r[3],
                entries: entries,
                key: getMaterialNameFromArtboardName(ab.name) + "|" + parsed.side + "|" + entries.length +
                     "|" + Math.round(r[2] - r[0]) + "x" + Math.round(r[1] - r[3])
            });
        }

        function sameLayout(a, b) {
            if (Math.abs(a.w - b.w) > TOL || Math.abs(a.h - b.h) > TOL) { return false; }
            if (a.entries.length !== b.entries.length) { return false; }
            for (var i = 0; i < a.entries.length; i++) {
                var x = a.entries[i];
                var y = b.entries[i];
                if (x.layer !== y.layer) { return false; }
                if (Math.abs(x.l - y.l) > TOL || Math.abs(x.t - y.t) > TOL ||
                    Math.abs(x.w - y.w) > TOL || Math.abs(x.h - y.h) > TOL) { return false; }
            }
            return true;
        }

        // Cluster boards with identical layouts
        var groups = [];
        var used = {};
        for (var i = 0; i < boards.length; i++) {
            if (used[i]) { continue; }
            var g = [boards[i]];
            for (var j = i + 1; j < boards.length; j++) {
                if (used[j] || boards[j].key !== boards[i].key) { continue; }
                if (sameLayout(boards[i], boards[j])) { g.push(boards[j]); used[j] = true; }
            }
            if (g.length > 1) { groups.push(g); }
        }
        if (groups.length === 0) { return result; }

        // Describe each group
        var baseName = originalDocPath.name.replace(/\.(ai|pdf|eps|svg)$/i, "").replace("_PRIME", "");
        for (var gi = 0; gi < groups.length; gi++) {
            var parts = [];
            for (var k = 0; k < groups[gi].length; k++) { parts.push(groups[gi][k].parsed.parts[0]); }
            var p0 = groups[gi][0].parsed;
            groups[gi].tail = ProdName.buildTail(p0.material, p0.blkbk, p0.side, parts, 1);
            var ptLabels = [];
            var sortedParts = parts.slice(0).sort(function (x, y) { return x - y; });
            for (var pl = 0; pl < sortedParts.length; pl++) { ptLabels.push("Pt" + sortedParts[pl]); }
            groups[gi].label = getMaterialNameFromArtboardName(doc.artboards[groups[gi][0].index].name) +
                (p0.side ? " " + p0.side : "") + ": " + ptLabels.join(", ") + "  ->  CUT_" + groups[gi].tail;
        }

        // One dialog for all groups
        var dlg = new Window("dialog", "Shared Cut Files");
        dlg.orientation = "column";
        dlg.alignChildren = "fill";
        dlg.margins = 20;
        dlg.spacing = 10;
        var head = dlg.add("statictext", undefined, "These artboards have identical cut layouts. Combine into one CUT file?");
        head.graphics.font = ScriptUI.newFont("dialog", "bold", 12);
        var boxes = [];
        for (var bi = 0; bi < groups.length; bi++) {
            var cb = dlg.add("checkbox", undefined, groups[bi].label);
            cb.value = true;
            boxes.push(cb);
        }
        dlg.add("statictext", undefined, "Unchecked groups keep a separate CUT file per artboard.");
        var btns = dlg.add("group");
        btns.alignment = "center";
        var combineBtn = btns.add("button", undefined, "Combine checked");
        var separateBtn = btns.add("button", undefined, "Keep all separate");
        var combine = false;
        combineBtn.onClick = function () { combine = true; dlg.close(); };
        separateBtn.onClick = function () { combine = false; dlg.close(); };
        dlg.show();
        if (!combine) { return result; }

        for (var gj = 0; gj < groups.length; gj++) {
            if (!boxes[gj].value) { continue; }
            // Lowest artboard index writes the shared CUT; the rest write none
            var first = groups[gj][0].index;
            for (var m = 1; m < groups[gj].length; m++) {
                if (groups[gj][m].index < first) { first = groups[gj][m].index; }
            }
            for (var m2 = 0; m2 < groups[gj].length; m2++) {
                var ix = groups[gj][m2].index;
                if (ix === first) { result.tail[ix] = groups[gj].tail; }
                else { result.skip[ix] = true; }
            }
        }
        return result;
    }

 
	// Check if a specific material uses CutContour by examining its artboards
	function checkMaterialUsesCutContour(materialName) {
		// Get the list of artboard indices for this material
		var materialInfo = materialData[materialName];
		if (!materialInfo || !materialInfo.artboards) {
			return false;
		}
		
		// Check each artboard for this material to see if any have CutContour content
		for (var i = 0; i < materialInfo.artboards.length; i++) {
			var artboardIndex = materialInfo.artboards[i].index;
			
			// Set active artboard to check its content
			var originalActiveIndex = doc.artboards.getActiveArtboardIndex();
			doc.artboards.setActiveArtboardIndex(artboardIndex);
			
			// Select all objects on this artboard
			doc.selectObjectsOnActiveArtboard();
			
			// Check if any selected objects are on CutContour layers
			for (var j = 0; j < doc.selection.length; j++) {
				var item = doc.selection[j];
				if (item.layer && (item.layer.name === "CutContour" || item.layer.name === "CutContour_PRF")) {
					// Found CutContour content for this material
					doc.selection = null; // Clear selection
					doc.artboards.setActiveArtboardIndex(originalActiveIndex); // Restore original artboard
					return true;
				}
			}
			
			// Clear selection and restore original artboard
			doc.selection = null;
			doc.artboards.setActiveArtboardIndex(originalActiveIndex);
		}
		
		return false; // No CutContour content found for this material
	}
 
    // Check if a material uses CutContour (which generates PnC files, not separate PRINT/CUT)
    function materialUsesCutContour(materialName) {
        // Find layers that might indicate CutContour usage
        for (var i = 0; i < doc.layers.length; i++) {
            var layer = doc.layers[i];
            if ((layer.name === "CutContour" || layer.name === "CutContour_PRF") && layerHasContent(layer)) {
                return true;
            }
        }
        return false;
    }
    
    // Check if filename contains PRIME
    function checkForPrime() {
        var fileName = doc.name;
        if (fileName.indexOf("PRIME") === -1) {
            alert("File name must contain 'PRIME'. Please save the file as a PRIME file before continuing.");
            return false;
        }
        return true;
    }
    
    // Check if artboards are properly named
    function checkArtboardNames() {
        var hasUnnamedArtboards = false;
        var unnamedList = [];
        
        for (var i = 0; i < doc.artboards.length; i++) {
            var artboardName = doc.artboards[i].name;
            var artboardNameLower = artboardName.toLowerCase();
            // Check if name is just a number, contains "artboard" (case-insensitive), or is "print" (case-insensitive)
            if (/^\d+$/.test(artboardName) || /artboard/i.test(artboardName) || artboardNameLower === "print") {
                hasUnnamedArtboards = true;
                unnamedList.push("Artboard " + (i + 1) + ": '" + artboardName + "'");
            }
        }
        
        if (hasUnnamedArtboards) {
            var message = "The following artboards are not properly named:\n" + 
                         unnamedList.join("\n") + 
                         "\n\nDo you want to continue anyway?";
            var firstConfirm = confirm(message);
            
            if (firstConfirm) {
                // Second warning
                var secondConfirm = confirm("Are you REALLY sure?\n\nIt is HIGHLY recommended that you cancel and correct file names first.\n\n'Whatever! I Do What I Want!'\n\nYes to do what you want. No to stop.");
                
                if (secondConfirm) {
                    // Third warning with random number verification
                    var randomNumber = Math.floor(Math.random() * 900000000) + 100000000; // Generate 9-digit random number
                    var userInput = prompt("Samir! You're breaking the car!\n\nVerification Number: " + randomNumber + "\n\nEnter the number above to continue or click Cancel to fix the artboard names:");
                    
                    if (userInput === null) {
                        return false; // User clicked Cancel
                    }
                    
                    if (userInput === randomNumber.toString()) {
                        return true; // User entered correct number
                    } else {
                        alert("Incorrect number entered.\n\nPlease fix your artboard names.");
                        return false;
                    }
                }
                return false;
            }
            return false;
        }
        return true;
    }
    
    // Check if artboards have illegal characters for file systems
    function checkArtboardCharacters() {
        var hasIllegalChars = false;
        var problemList = [];
        // Using string method instead of regex to avoid escape issues
        var illegalCharsString = '<>:"/\\|?*';
        
        for (var i = 0; i < doc.artboards.length; i++) {
            var artboardName = doc.artboards[i].name;
            var foundChars = [];
            
            // Check each illegal character
            for (var j = 0; j < illegalCharsString.length; j++) {
                if (artboardName.indexOf(illegalCharsString[j]) !== -1) {
                    if (foundChars.indexOf(illegalCharsString[j]) === -1) {
                        foundChars.push(illegalCharsString[j]);
                    }
                }
            }
            
            if (foundChars.length > 0) {
                hasIllegalChars = true;
                problemList.push("Artboard " + (i + 1) + " '" + artboardName + "' contains: " + foundChars.join(" "));
            }
        }
        
        if (hasIllegalChars) {
            alert("The following artboards have illegal characters for file names:\n\n" + 
                  problemList.join("\n") + 
                  "\n\nPlease rename these artboards to remove illegal characters.");
            return false;
        }
        return true;
    }
    
    // Process artboards - selectedIndices is an array of artboard index numbers to export
    function processAllArtboards(selectedIndices) {
        var indicesToProcess = (selectedIndices && selectedIndices.length > 0) ? selectedIndices : [];

        // Fallback: if nothing passed in, process everything
        if (indicesToProcess.length === 0) {
            for (var x = 0; x < doc.artboards.length; x++) indicesToProcess.push(x);
        }

        for (var ii = 0; ii < indicesToProcess.length; ii++) {
            var i = indicesToProcess[ii];
            var currentStep = "starting";
            var artboardName = "(unknown)";
            try {
                // Set active artboard
                currentStep = "setActiveArtboardIndex(" + i + ")";
                doc.artboards.setActiveArtboardIndex(i);

                currentStep = "reading artboard name at index " + i;
                artboardName = doc.artboards[i].name;

                // Isolate current artboard
                currentStep = "isolateArtboard(" + i + ")";
                var hasContent = isolateArtboard(i);

                // Skip if artboard is empty
                if (!hasContent) {
                    continue;
                }

                // Force completion of isolation
                currentStep = "app.redraw()";
                app.redraw();

                // Save isolated state to temp file for duplication
                currentStep = "saving temp file";
                var tempPath = Folder.temp + "/isolated_" + Date.now() + ".ai";
                var tempFile = new File(tempPath);
                var aiOptions = new IllustratorSaveOptions();
                aiOptions.compatibility = Compatibility.ILLUSTRATOR24;
                doc.saveAs(tempFile, aiOptions);

                // Check what type of file to create
                // Check if CutContour or CutContour_PRF layers have content on this artboard
                currentStep = "getCutContourInfo()";
                var cutContourInfo = getCutContourInfo();

                if (cutContourInfo.hasCutContour) {
                    if (cutContourInfo.hasPrintData) {
                        currentStep = "createPnCFileFromTemp";
                        createPnCFileFromTemp(tempFile, artboardName);
                    } else {
                        currentStep = "createCutOnlyFileFromTemp";
                        createCutOnlyFileFromTemp(tempFile, artboardName);
                    }
                } else {
                    currentStep = "createPrintFileFromTemp";
                    createPrintFileFromTemp(tempFile, artboardName);
                    currentStep = "createCutFileFromTemp";
                    // Shared cut layouts: one artboard writes CUT_..._Pt3-5, the rest write none
                    if (!sharedCuts.skip[i]) {
                        createCutFileFromTemp(tempFile, sharedCuts.tail[i] || artboardName);
                    }
                }

                // Delete temp file
                currentStep = "tempFile.remove()";
                tempFile.remove();

                // Reopen original document
                currentStep = "app.open(originalDocPath)";
                doc = app.open(originalDocPath);

                // Restore scale factor reference after reopening
                scaleFactor = doc.scaleFactor || 1;

            } catch (e) {
                var docName = "(unknown)";
                try { docName = app.activeDocument.name; } catch (eDoc) {}

                var details = "Error processing artboard " + (i + 1) + " (" + artboardName + ")\n\n" +
                              "Failed step: " + currentStep + "\n" +
                              "Active document: " + docName + "\n" +
                              "Error: " + e.toString() + "\n" +
                              "Line: " + (e.line !== undefined ? e.line : "(unknown)") + "\n" +
                              (e.fileName ? "File: " + e.fileName + "\n" : "") +
                              "\nStopping process.";

                $.writeln(details);
                alert(details);
                throw e;
            }
        }
    }
    
// Isolate current artboard
function isolateArtboard(artboardIndex) {
    // Store the current artboard's bounds for reference
    var currentArtboard = doc.artboards[artboardIndex];
    var artboardBounds = currentArtboard.artboardRect;
    
    // Select all on active artboard
    doc.artboards.setActiveArtboardIndex(artboardIndex);
    doc.selectObjectsOnActiveArtboard();
    
    // If nothing selected on this artboard, skip it
    if (doc.selection.length === 0) {
        return false; // Return false to indicate empty artboard
    }
    
    // Select inverse and delete
    app.executeMenuCommand("Inverse menu item");
    
    // Delete inverse selection
    if (doc.selection.length > 0) {
        for (var j = doc.selection.length - 1; j >= 0; j--) {
            doc.selection[j].remove();
        }
    }
    
    // Clear selection
    doc.selection = null;
    
    // Remove all other artboards but keep the current one
    var newIndex = 0;
    for (var k = doc.artboards.length - 1; k >= 0; k--) {
        if (k !== artboardIndex) {
            doc.artboards.remove(k);
            if (k < artboardIndex) {
                newIndex = artboardIndex - 1;
            } else {
                newIndex = artboardIndex;
            }
            artboardIndex = newIndex;
        }
    }
    
    doc.artboards.setActiveArtboardIndex(0);
    doc.selection = null;
    
    return true;
}
    
	// Get information about CutContour layers and print data
	function getCutContourInfo() {
		var cutContourLayer = null;
		var cutContourPrfLayer = null;
		var printLayers = []; // Changed to array to handle multiple print layer types
		
		// Find relevant layers
		for (var i = 0; i < doc.layers.length; i++) {
			var layerName = doc.layers[i].name;
			var layerNameLower = layerName.toLowerCase();
			
			if (layerName === "CutContour") {
				cutContourLayer = doc.layers[i];
			} else if (layerName === "CutContour_PRF") {
				cutContourPrfLayer = doc.layers[i];
			} else if (layerNameLower === "print" || layerNameLower === "spot1") {
				// Now recognizes both "print" and "spot1" as print layers
				printLayers.push(doc.layers[i]);
			}
		}
		
		// Check if any CutContour layers have content
		var hasCutContour = false;
		if (cutContourLayer && layerHasContent(cutContourLayer)) {
			hasCutContour = true;
		}
		if (cutContourPrfLayer && layerHasContent(cutContourPrfLayer)) {
			hasCutContour = true;
		}
		
		// Check if any print layers have content
		var hasPrintData = false;
		for (var j = 0; j < printLayers.length; j++) {
			if (layerHasContent(printLayers[j])) {
				hasPrintData = true;
				break; // Found at least one print layer with content
			}
		}
		
		return {
			hasCutContour: hasCutContour,
			hasPrintData: hasPrintData
		};
    }
    
    // Create PRINT file from temp file
    function createPrintFileFromTemp(tempFile, artboardName) {
        // PRESERVE CONTEXT before opening temp file
        var savedMaterialData = materialData;
        var savedUseFolderOrganization = useFolderOrganization;
        var savedParentSavePath = parentSavePath;
        var savedOverwriteAll = overwriteAll;
        var savedSkipAll = skipAll;
        
        // Open the temp file
        var tempDoc = app.open(tempFile);
        
        var hasAnyContent = false;
        
        // Delete layers we don't want for PRINT file
        for (var i = tempDoc.layers.length - 1; i >= 0; i--) {
            var layer = tempDoc.layers[i];
            var layerName = layer.name.toLowerCase();
            
            // Delete cut layers (including CutContour and CutContour_PRF)
            if (layerName.indexOf("cut") !== -1) {
                layer.remove();
                continue;
            }
            
            // Keep only print/reg/spot1 that have content
            if (layerName === "print" || layerName === "reg" || layerName === "spot1") {
                if (!layerHasContent(layer)) {
                    layer.remove();
                } else {
                    hasAnyContent = true;
                }
            } else {
                layer.remove();
            }
        }
        
        // Check if document has any content left
        if (!hasAnyContent || tempDoc.layers.length === 0) {
            // Skip this file - no content
            tempDoc.close(SaveOptions.DONOTSAVECHANGES);
            return;
        }
        
        // Save as PDF with preserved context
        var fileName = generateFileName("PRINT", artboardName);
        saveFileAsPDFWithContext(tempDoc, fileName, artboardName, 
                                savedMaterialData, savedUseFolderOrganization, 
                                savedParentSavePath, savedOverwriteAll, savedSkipAll);
        
        // Close without saving
        tempDoc.close(SaveOptions.DONOTSAVECHANGES);
    }
    
    // Create CUT file from temp file
    function createCutFileFromTemp(tempFile, artboardName) {
        // PRESERVE CONTEXT before opening temp file
        var savedMaterialData = materialData;
        var savedUseFolderOrganization = useFolderOrganization;
        var savedParentSavePath = parentSavePath;
        var savedOverwriteAll = overwriteAll;
        var savedSkipAll = skipAll;
        
        // Open the temp file
        var tempDoc = app.open(tempFile);
        
        var hasKnifeCut = false;
        var knifeCutLayerName = "";
        var otherCuts = [];
        var hasAnyContent = false;
        
        // Delete and check layers
        for (var i = tempDoc.layers.length - 1; i >= 0; i--) {
            var layer = tempDoc.layers[i];
            var layerName = layer.name;
            var layerNameLower = layerName.toLowerCase();
            
            // Delete PRINT and Spot1
            if (layerNameLower === "print" || layerNameLower === "spot1") {
                layer.remove();
                continue;
            }
            
            // Delete CutContour and CutContour_PRF layers (they should go to PnC or separate files)
            if (layerName === "CutContour" || layerName === "CutContour_PRF") {
                layer.remove();
                continue;
            }
            
            // Keep REG if it has content
            if (layerNameLower === "reg") {
                if (!layerHasContent(layer)) {
                    layer.remove();
                } else {
                    hasAnyContent = true;
                }
                continue;
            }
            
            // Keep pin hole layers if they have content.
            // Matches both the legacy "PinMountHoles-Inside" name and the
            // current "PinHoles" name (and any other pin/hole variant).
            if (isPinHoleLayerName(layerName)) {
                if (!layerHasContent(layer)) {
                    layer.remove();
                } else {
                    hasAnyContent = true;
                    otherCuts.push(layerName);
                }
                continue;
            }
            
            // Process other cut layers (excluding CutContour types)
            if (layerNameLower.indexOf("cut") !== -1) {
                if (!layerHasContent(layer)) {
                    layer.remove();
                } else {
                    hasAnyContent = true;
                    if (isKnifeCutLayerName(layerName)) {
                        hasKnifeCut = true;
                        knifeCutLayerName = layerName;
                    } else {
                        otherCuts.push(layerName);
                    }
                }
            } else {
                layer.remove();
            }
        }
        
        // Check if document has any content left
        if (!hasAnyContent || tempDoc.layers.length === 0) {
            // Skip this file - no content
            tempDoc.close(SaveOptions.DONOTSAVECHANGES);
            return;
        }
        
        // Warn if knife cut is with other cuts
        if (hasKnifeCut && otherCuts.length > 0) {
            alert("Warning: Knife cut (" + knifeCutLayerName + ") is on the same artboard as other cut types:\n" + 
                  otherCuts.join(", "));
        }
        
        // Save as PDF with preserved context
        var fileName = generateFileName("CUT", artboardName);
        saveFileAsPDFWithContext(tempDoc, fileName, artboardName, 
                                savedMaterialData, savedUseFolderOrganization, 
                                savedParentSavePath, savedOverwriteAll, savedSkipAll);
        
        // Close without saving
        tempDoc.close(SaveOptions.DONOTSAVECHANGES);
    }
    
    // Create PnC file from temp file (when CutContour/CutContour_PRF + Print data present)
    function createPnCFileFromTemp(tempFile, artboardName) {
        // PRESERVE CONTEXT before opening temp file
        var savedMaterialData = materialData;
        var savedUseFolderOrganization = useFolderOrganization;
        var savedParentSavePath = parentSavePath;
        var savedOverwriteAll = overwriteAll;
        var savedSkipAll = skipAll;
        
        // Open the temp file
        var tempDoc = app.open(tempFile);
        
        var hasAnyContent = false;
        
        // Delete layers we don't want
        for (var i = tempDoc.layers.length - 1; i >= 0; i--) {
            var layer = tempDoc.layers[i];
            var layerName = layer.name;
            var layerNameLower = layerName.toLowerCase();
            
            // Keep only print/spot1/cutcontour/cutcontour_prf that have content
            if (layerNameLower === "print" || layerNameLower === "spot1" || 
                layerName === "CutContour" || layerName === "CutContour_PRF") {
                if (!layerHasContent(layer)) {
                    layer.remove();
                } else {
                    hasAnyContent = true;
                }
            } else {
                layer.remove();
            }
        }
        
        // Check if document has any content left
        if (!hasAnyContent || tempDoc.layers.length === 0) {
            // Skip this file - no content
            tempDoc.close(SaveOptions.DONOTSAVECHANGES);
            return;
        }
        
        // Save as PDF with preserved context
        var fileName = generateFileName("PnC", artboardName);
        saveFileAsPDFWithContext(tempDoc, fileName, artboardName, 
                                savedMaterialData, savedUseFolderOrganization, 
                                savedParentSavePath, savedOverwriteAll, savedSkipAll);
        
        // Close without saving
        tempDoc.close(SaveOptions.DONOTSAVECHANGES);
    }
    
    // Create CutOnly file from temp file (when CutContour/CutContour_PRF present but no Print data)
    function createCutOnlyFileFromTemp(tempFile, artboardName) {
        // PRESERVE CONTEXT before opening temp file
        var savedMaterialData = materialData;
        var savedUseFolderOrganization = useFolderOrganization;
        var savedParentSavePath = parentSavePath;
        var savedOverwriteAll = overwriteAll;
        var savedSkipAll = skipAll;
        
        // Open the temp file
        var tempDoc = app.open(tempFile);
        
        var hasAnyContent = false;
        
        // Delete layers we don't want
        for (var i = tempDoc.layers.length - 1; i >= 0; i--) {
            var layer = tempDoc.layers[i];
            var layerName = layer.name;
            var layerNameLower = layerName.toLowerCase();
            
            // Keep only cutcontour/cutcontour_prf/reg/pin holes that have content
            if (layerName === "CutContour" || layerName === "CutContour_PRF" ||
                layerNameLower === "reg" || isPinHoleLayerName(layerName)) {
                if (!layerHasContent(layer)) {
                    layer.remove();
                } else {
                    hasAnyContent = true;
                }
            } else {
                layer.remove();
            }
        }
        
        // Check if document has any content left
        if (!hasAnyContent || tempDoc.layers.length === 0) {
            // Skip this file - no content
            tempDoc.close(SaveOptions.DONOTSAVECHANGES);
            return;
        }
        
        // Save as PDF with preserved context
        var fileName = generateFileName("CutOnly", artboardName);
        saveFileAsPDFWithContext(tempDoc, fileName, artboardName, 
                                savedMaterialData, savedUseFolderOrganization, 
                                savedParentSavePath, savedOverwriteAll, savedSkipAll);
        
        // Close without saving
        tempDoc.close(SaveOptions.DONOTSAVECHANGES);
    }
    
    // True for any pin-mount hole layer, regardless of naming convention
    // ("PinHoles", "PinMountHoles-Inside", "Pin Mount Holes", etc.)
    function isPinHoleLayerName(layerName) {
        var n = String(layerName).toLowerCase().replace(/[^a-z]/g, "");
        return n.indexOf("pin") !== -1 && n.indexOf("hole") !== -1;
    }

    // True for any knife cut layer, regardless of naming convention
    // ("CutThrough-Knifecut", "CutThrough3-Knifecut", "Knife Cut", etc.)
    function isKnifeCutLayerName(layerName) {
        var n = String(layerName).toLowerCase().replace(/[^a-z]/g, "");
        return n.indexOf("knifecut") !== -1;
    }

    // Check if layer has content
    function layerHasContent(layer) {
        if (layer.pageItems.length > 0) {
            return true;
        }
        
        // Check sublayers
        for (var i = 0; i < layer.layers.length; i++) {
            if (layerHasContent(layer.layers[i])) {
                return true;
            }
        }
        
        return false;
    }
    
    // New saveFileAsPDFWithContext function that accepts preserved context
    function saveFileAsPDFWithContext(tempDoc, fileName, artboardName, 
                                     contextMaterialData, contextUseFolderOrganization, 
                                     contextParentSavePath, contextOverwriteAll, contextSkipAll) {
        var saveFile;
        
        // Use the preserved context instead of global variables
        if (contextUseFolderOrganization && artboardName) {
            var materialName = getMaterialNameFromArtboardName(artboardName);
            var materialFolder = createMaterialFolderWithContext(materialName, contextParentSavePath);
            
            if (!materialFolder) {
                // Folder creation failed, ask user what to do
                var fallbackChoice = confirm("Failed to create folder for material '" + materialName + "'.\n\n" +
                                           "Click OK to save in the main directory, or Cancel to stop processing.");
                if (!fallbackChoice) {
                    return false; // Stop processing
                }
                // Continue with main directory
                saveFile = new File(contextParentSavePath + "/" + fileName);
            } else {
                // Check if THIS SPECIFIC material needs subfolders using preserved context
                var materialInfo = contextMaterialData[materialName];
                var needsSubfolders = false;
                
                // Only create subfolders if this material has >10 parts
                // (We can't check CutContour from temp doc context, so simplified logic)
                if (materialInfo && materialInfo.artboards.length > 10) {
                    needsSubfolders = true;
                }
                
                if (needsSubfolders) {
                    // Create subfolders for PRINT and CUT files only
                    if (fileName.indexOf("_PRINT_") !== -1) {
                        var printFolder = createSubFolder(materialFolder, "PRINT");
                        saveFile = new File((printFolder || materialFolder) + "/" + fileName);
                    } else if (fileName.indexOf("_CUT_") !== -1) {
                        var cutFolder = createSubFolder(materialFolder, "CUT");
                        saveFile = new File((cutFolder || materialFolder) + "/" + fileName);
                    } else {
                        // PnC or CutOnly files go in material root folder
                        saveFile = new File(materialFolder + "/" + fileName);
                    }
                } else {
                    // This material has <=10 parts, no subfolders needed
                    saveFile = new File(materialFolder + "/" + fileName);
                }
            }
        } else {
            // Use the preserved parent save path
            saveFile = new File(contextParentSavePath + "/" + fileName);
        }
        
        // Check if file exists
        if (saveFile.exists) {
            // Check skipAll first - if true, always skip
            if (contextSkipAll) {
                return; // Skip this file
            }
            
            var overwrite = false;
            
            if (contextOverwriteAll === null) {
                var dialog = confirm("File '" + fileName + "' already exists.\nDo you want to overwrite it?\n\n" +
                                   "Click 'Yes' to overwrite, 'No' to skip.");
                
                if (dialog) {
                    overwrite = true;
                    // Ask if should apply to all
                    if (confirm("Apply this decision to all existing files?")) {
                        // Update global variable
                        overwriteAll = true;
                    }
                } else {
                    if (confirm("Skip all existing files?")) {
                        // Update global variable
                        skipAll = true;
                        return; // Skip this file immediately
                    }
                }
            } else {
                overwrite = contextOverwriteAll;
            }
            
            if (!overwrite) {
                return; // Skip this file
            }
        }
        
        // Set up PDF save options with Large Canvas compatibility
        var pdfOptions = new PDFSaveOptions();
        pdfOptions.compatibility = PDFCompatibility.ACROBAT7; // PDF 1.6 for Large Canvas support
        pdfOptions.generateThumbnails = true;
        pdfOptions.preserveEditability = true;
        pdfOptions.preset = "[Illustrator Default]";
        pdfOptions.viewAfterSaving = false;
        
        // Save as PDF directly
        tempDoc.saveAs(saveFile, pdfOptions);
        
        // Force completion
        app.redraw();
    }
    
    // Helper function to create material folder with specific parent path
    function createMaterialFolderWithContext(materialName, contextParentSavePath) {
        try {
            var materialFolder = new Folder(contextParentSavePath + "/" + materialName);
            if (!materialFolder.exists) {
                materialFolder.create();
            }
            return materialFolder.fsName;
        } catch (e) {
            return null; // Folder creation failed
        }
    }
    
    // Save document as PDF with Large Canvas scaling fix and folder organization (LEGACY - kept for compatibility)
	function saveFileAsPDF(tempDoc, fileName, artboardName) {
		var saveFile;
		
		// Determine save path based on folder organization setting
		if (useFolderOrganization && artboardName) {
			var materialName = getMaterialNameFromArtboardName(artboardName);
			var materialFolder = createMaterialFolder(materialName);
			
			if (!materialFolder) {
				// Folder creation failed, ask user what to do
				var fallbackChoice = confirm("Failed to create folder for material '" + materialName + "'.\n\n" +
										   "Click OK to save in the main directory, or Cancel to stop processing.");
				if (!fallbackChoice) {
					return false; // Stop processing
				}
				// Continue with main directory
				saveFile = new File(parentSavePath + "/" + fileName);
			} else {
				// Check if THIS SPECIFIC material needs subfolders 
				// (>10 artboards for this material AND not using CutContour)
				var materialInfo = materialData[materialName];
				var needsSubfolders = false;
				
				// Only create subfolders if this material has >10 parts AND doesn't use CutContour
				if (materialInfo && materialInfo.artboards.length > 10) {
					// Check if this specific material uses CutContour by checking its artboards
					var materialUsesCutContourFlag = checkMaterialUsesCutContour(materialName);
					if (!materialUsesCutContourFlag) {
						needsSubfolders = true;
					}
				}
				
				if (needsSubfolders) {
					// Create subfolders for PRINT and CUT files only
					if (fileName.indexOf("_PRINT_") !== -1) {
						var printFolder = createSubFolder(materialFolder, "PRINT");
						saveFile = new File((printFolder || materialFolder) + "/" + fileName);
					} else if (fileName.indexOf("_CUT_") !== -1) {
						var cutFolder = createSubFolder(materialFolder, "CUT");
						saveFile = new File((cutFolder || materialFolder) + "/" + fileName);
					} else {
						// PnC or CutOnly files go in material root folder
						saveFile = new File(materialFolder + "/" + fileName);
					}
				} else {
					// This material has <=10 parts or uses CutContour, no subfolders needed
					// ALL files go directly in the material folder
					saveFile = new File(materialFolder + "/" + fileName);
				}
			}
		} else {
			// Use the stored parent save path
			saveFile = new File(parentSavePath + "/" + fileName);
		}
		
		// Check if file exists
        if (saveFile.exists) {
            // Check skipAll first - if true, always skip
            if (skipAll) {
                return; // Skip this file
            }
            
            var overwrite = false;
            
            if (overwriteAll === null) {
                var dialog = confirm("File '" + fileName + "' already exists.\nDo you want to overwrite it?\n\n" +
                                   "Click 'Yes' to overwrite, 'No' to skip.");
                
                if (dialog) {
                    overwrite = true;
                    // Ask if should apply to all
                    if (confirm("Apply this decision to all existing files?")) {
                        overwriteAll = true;
                    }
                } else {
                    if (confirm("Skip all existing files?")) {
                        skipAll = true;
                        return; // Skip this file immediately
                    }
                }
            } else {
                overwrite = overwriteAll;
            }
            
            if (!overwrite) {
                return; // Skip this file
            }
        }
		
		// Set up PDF save options with Large Canvas compatibility
		var pdfOptions = new PDFSaveOptions();
		pdfOptions.compatibility = PDFCompatibility.ACROBAT7; // PDF 1.6 for Large Canvas support
		pdfOptions.generateThumbnails = true;
		pdfOptions.preserveEditability = true;
		pdfOptions.preset = "[Illustrator Default]";
		pdfOptions.viewAfterSaving = false;
		
		// Save as PDF directly
		tempDoc.saveAs(saveFile, pdfOptions);
		
		// Force completion
		app.redraw();
	}
    
    // Extract artboard name from filename for folder organization
    function getArtboardNameFromFileName(fileName) {
        // Format: basename_TYPE_artboardname.pdf
        var parts = fileName.split('_');
        if (parts.length >= 3) {
            var artboardName = parts.slice(2).join('_').replace('.pdf', '');
            return artboardName;
        }
        return "Unknown";
    }
    
    // Extract material name from artboard name for folder organization
    function getMaterialNameFromArtboardName(artboardName) {
        // Material group used for folders: material plus _BlkBk when on black-back
        // stock (parts, side and _x# stripped). "3mmACM_BlkBk_Pt2_x2" -> "3mmACM_BlkBk"
        var p = ProdName.parseArtboardName(String(artboardName).replace(/^\s+|\s+$/g, ''));
        if (!p.material) { return artboardName; }
        return ProdName.buildTail(p.material, p.blkbk, '', [], 1);
    }

    // Standard file tail for an artboard name (legacy names normalized):
    // "3mmACMBLK_Pt1" -> "3mmACM_BlkBk_Pt1", "3mmACM_Pt7_x3" unchanged
    function standardTail(artboardName, forCut) {
        var p = ProdName.parseArtboardName(String(artboardName).replace(/^\s+|\s+$/g, ''));
        if (!p.material) { return artboardName; }
        var side = p.side;
        // CUTs only come from FRONT sheets, so CUT_3mmACM_Pt14 is enough - unless an
        // older PRIME also has a one-sided 3mmACM_Pt14 (then keep _FRONT so the
        // two CUT files can't overwrite each other).
        if (forCut && side === 'FRONT' && !oneSidedPartExists(p)) { side = ''; }
        return ProdName.buildTail(p.material, p.blkbk, side, p.parts, p.mult);
    }

    function oneSidedPartExists(fp) {
        try {
            var want = ProdName.matKey(ProdName.buildTail(fp.material, fp.blkbk, '', [], 1));
            var wantPt = fp.parts.length ? fp.parts[0] : 1;
            for (var i = 0; i < doc.artboards.length; i++) {
                var q = ProdName.parseArtboardName(doc.artboards[i].name);
                if (!q.material || q.side) { continue; }
                if (ProdName.matKey(ProdName.buildTail(q.material, q.blkbk, '', [], 1)) !== want) { continue; }
                var qp = q.parts.length ? q.parts : [1];
                for (var k = 0; k < qp.length; k++) { if (qp[k] === wantPt) { return true; } }
            }
        } catch (e) { return true; } // unsure: keep FRONT (never risk a name clash)
        return false;
    }
    
    // Create material folder and return path
    function createMaterialFolder(materialName) {
        try {
            // Use the original case from materialData if available
            var folderName = materialName;
            if (materialData[materialName] && materialData[materialName].originalCase) {
                folderName = materialData[materialName].originalCase;
            }
            
            var materialFolder = new Folder(parentSavePath + "/" + folderName);
            if (!materialFolder.exists) {
                materialFolder.create();
            }
            return materialFolder.fsName;
        } catch (e) {
            return null; // Folder creation failed
        }
    }
    
    // Create subfolder (PRINT or CUT) and return path
    function createSubFolder(parentFolderPath, subFolderName) {
        try {
            var subFolder = new Folder(parentFolderPath + "/" + subFolderName);
            if (!subFolder.exists) {
                subFolder.create();
            }
            return subFolder.fsName;
        } catch (e) {
            return null; // Subfolder creation failed
        }
    }
    
    // Generate filename with type prefix
    function generateFileName(type, artboardName) {
        // Use the stored original document path to get the name
        var originalName = originalDocPath.name;
        // Remove any file extension (.ai, .pdf, .eps, .svg)
        var baseName = originalName.replace(/\.(ai|pdf|eps|svg)$/i, "");
        baseName = baseName.replace("_PRIME", "");
        
        return baseName + "_" + type + "_" + standardTail(artboardName, type === "CUT") + ".pdf";
    }
    
})();