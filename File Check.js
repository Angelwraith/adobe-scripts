#target illustrator

/*@METADATA{
  "name": "File Check",
  "description": "Production file readiness analysis with manifest comparison",
  "version": "3.2",
  "target": "illustrator",
  "tags": ["file", "check", "report", "manifest"]
}@END_METADATA*/

// JSON Polyfill for ExtendScript (doesn't have native JSON)
if (typeof JSON === 'undefined') {
    var JSON = {
        stringify: function(obj) {
            if (obj === null) return 'null';
            if (obj === undefined) return undefined;
            if (typeof obj === 'string') return '"' + obj.replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/\n/g, '\\n').replace(/\r/g, '\\r') + '"';
            if (typeof obj === 'number') return isFinite(obj) ? String(obj) : 'null';
            if (typeof obj === 'boolean') return String(obj);
            if (obj instanceof Array) {
                var items = [];
                for (var i = 0; i < obj.length; i++) {
                    items.push(this.stringify(obj[i]));
                }
                return '[' + items.join(',') + ']';
            }
            if (typeof obj === 'object') {
                var items = [];
                for (var key in obj) {
                    if (obj.hasOwnProperty(key)) {
                        items.push('"' + key + '":' + this.stringify(obj[key]));
                    }
                }
                return '{' + items.join(',') + '}';
            }
            return '""';
        },
        parse: function(str) {
            return eval('(' + str + ')');
        }
    };
}

// Main entry point with proper error checking
try {
    if (app.documents.length == 0) {
        alert("Please open at least one document first.");
    } else {
        showDocumentSelectionDialog();
    }
} catch (e) {
    alert("Startup error: " + e.toString());
}

function showDocumentSelectionDialog() {
    var dialog = new Window("dialog", "File Analysis - Select Documents");
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.spacing = 15;
    dialog.margins = 20;
    dialog.preferredSize.width = 500;
    dialog.preferredSize.height = 175;

    var titleText = dialog.add("statictext", undefined, "Select documents to analyze:");
    titleText.graphics.font = ScriptUI.newFont("dialog", "Bold", 14);

    dialog.add("panel");

    // Document selection - NO Analysis Options panel
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

    // THREE BUTTONS - No radio buttons anywhere
    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "center";
    buttonGroup.spacing = 10;

    var cancelButton = buttonGroup.add("button", undefined, "Cancel");
    var quickButton = buttonGroup.add("button", undefined, "Quick Analysis");
    var ppiButton = buttonGroup.add("button", undefined, "Full Analysis (with PPI)");

    var result = null;

    cancelButton.onClick = function() {
        result = null;
        dialog.close();
    };

    quickButton.onClick = function() {
        var selectedDocuments = [];
        for (var i = 0; i < checkboxes.length; i++) {
            if (checkboxes[i].checkbox.value) {
                selectedDocuments.push(checkboxes[i].document);
            }
        }

        if (selectedDocuments.length === 0) {
            alert("Please select at least one document to analyze.");
            return;
        }

        result = {
            documents: selectedDocuments,
            includePPI: false
        };
        dialog.close();
    };

    ppiButton.onClick = function() {
        var selectedDocuments = [];
        for (var i = 0; i < checkboxes.length; i++) {
            if (checkboxes[i].checkbox.value) {
                selectedDocuments.push(checkboxes[i].document);
            }
        }

        if (selectedDocuments.length === 0) {
            alert("Please select at least one document to analyze.");
            return;
        }

        result = {
            documents: selectedDocuments,
            includePPI: true
        };
        dialog.close();
    };

    dialog.show();

    if (result && result.documents && result.documents.length > 0) {
        runAnalysis(result.documents, result.includePPI);
    }
}

function extractMaterialFromFilename(filename) {
    // Production file naming convention (from Optimal PRIME transformer):
    //   {baseName}_{type}_{artboardName}.pdf
    // type = PRINT, CUT, PnC, or CutOnly
    // artboardName = material name, optionally suffixed with _Pt# for multi-part
    var upperName = filename.toUpperCase();
    // Check longest keywords first to avoid partial matches (e.g. _CUT_ inside _CUTONLY_)
    var keywords = ['_CUTONLY_', '_PRINT_', '_PNC_', '_CUT_'];

    for (var k = 0; k < keywords.length; k++) {
        var idx = upperName.indexOf(keywords[k]);
        if (idx !== -1) {
            var afterKeyword = filename.substring(idx + keywords[k].length);
            // Remove file extension
            var extensionMatch = afterKeyword.match(/\.(ai|eps|pdf)$/i);
            if (extensionMatch) {
                afterKeyword = afterKeyword.substring(0, afterKeyword.length - extensionMatch[0].length);
            }
            // Strip _Pt# suffix (multi-part artboards) to group all parts under same material
            afterKeyword = afterKeyword.replace(/_Pt\d+$/i, '');
            return afterKeyword;
        }
    }
    return '';
}

function extractMaterialFromArtboards(doc) {
    // Fallback material detection for combined/PRIME files whose filename has no
    // _PRINT_/_CUT_/_PNC_/_CUTONLY_ token. Artboards in these files are named after
    // the material (e.g. "3mmACM_Pt1", "3mmACM_Pt2"). Strip the _Pt# suffix and, if
    // every artboard shares the same base name, use it as the material name.
    // If artboards disagree (mixed materials in one file), return '' rather than guess.
    try {
        if (!doc || !doc.artboards || doc.artboards.length === 0) {
            return '';
        }
        var firstMat = null;
        for (var a = 0; a < doc.artboards.length; a++) {
            var abName = doc.artboards[a].name;
            if (!abName) { continue; }
            // Strip trailing _Pt# (and _Part#) multi-part suffix to match the
            // normalization used by extractMaterialFromFilename
            var base = abName.replace(/_Pt\d+$/i, '').replace(/_Part\d+$/i, '');
            // Trim surrounding whitespace
            base = base.replace(/^\s+|\s+$/g, '');
            if (base.length === 0) { continue; }
            if (firstMat === null) {
                firstMat = base;
            } else if (firstMat !== base) {
                // Mixed materials across artboards - can't safely pick one
                return '';
            }
        }
        return firstMat || '';
    } catch (e) {
        return '';
    }
}

function extractFileType(filename) {
    // Returns the production file type: PRINT, CUT, CutOnly, PnC, or empty
    var upperName = filename.toUpperCase();
    if (upperName.indexOf('_CUTONLY_') !== -1) { return 'CutOnly'; }
    if (upperName.indexOf('_PNC_') !== -1) { return 'PnC'; }
    if (upperName.indexOf('_PRINT_') !== -1) { return 'PRINT'; }
    if (upperName.indexOf('_CUT_') !== -1) { return 'CUT'; }
    return '';
}

function normalizeDimension(width, height) {
    // Round to nearest 1/8"
    width = Math.round(width * 8) / 8;
    height = Math.round(height * 8) / 8;

    // Ensure smaller dimension first
    if (width <= height) {
        return width + '"x' + height + '"';
    } else {
        return height + '"x' + width + '"';
    }
}

function findAndReadProofManifest(documents) {
    // Returns { data: <object>|null, status: <string> }
    // status is always set so we can show diagnostics when manifest is not found
    try {
        if (!documents || documents.length === 0) {
            return { data: null, status: 'No documents provided' };
        }

        var firstDocPath = documents[0].fullName;
        if (!firstDocPath) {
            return { data: null, status: 'First document has no file path (unsaved?)' };
        }

        var docFile = new File(firstDocPath);
        var currentFolder = docFile.parent;
        var searchedFolders = [];

        // Walk up the folder hierarchy looking for a PRIME subfolder
        // Production files live in the order folder (or material subfolders within it)
        // PRIME folder is a child of the order folder:
        //   Order Folder/PRIME/proof.pdf
        //   Order Folder/production_file.pdf
        //   Order Folder/Material Subfolder/production_file.pdf
        var maxLevels = 5;
        var primeFolder = null;

        for (var level = 0; level < maxLevels; level++) {
            if (!currentFolder || currentFolder.parent === currentFolder) {
                break;
            }

            searchedFolders.push(decodeURI(currentFolder.name));

            // Check if PRIME folder exists as a child of current folder
            var primeFolderPath = currentFolder.absoluteURI + '/' + 'PRIME';
            var potentialPrimeFolder = new Folder(primeFolderPath);

            if (potentialPrimeFolder.exists) {
                primeFolder = potentialPrimeFolder;
                break;
            }

            currentFolder = currentFolder.parent;
        }

        if (!primeFolder) {
            return { data: null, status: 'No PRIME folder found. Searched: ' + searchedFolders.join(' > ') };
        }

        // Look for a file with "Proof" in the name (case-insensitive), preferring PDF
        var proofFile = null;
        var filesInPrime = [];
        var files = primeFolder.getFiles();

        for (var f = 0; f < files.length; f++) {
            var file = files[f];
            if (file instanceof File) {
                filesInPrime.push(file.name);
                var fileName = file.name.toUpperCase();
                if (fileName.indexOf('PROOF') !== -1) {
                    // Prefer PDF files
                    if (fileName.indexOf('.PDF') !== -1) {
                        proofFile = file;
                        break;
                    } else if (!proofFile) {
                        proofFile = file;
                    }
                }
            }
        }

        if (!proofFile) {
            return { data: null, status: 'PRIME folder found but no proof file. Files in PRIME: ' + filesInPrime.join(', ') };
        }

        // Read XMP metadata from proof file
        if (ExternalObject.AdobeXMPScript === undefined) {
            ExternalObject.AdobeXMPScript = new ExternalObject('lib:AdobeXMPScript');
        }

        // Determine file type for XMPFile constructor
        var proofNameUpper = proofFile.name.toUpperCase();
        var xmpFileType = XMPConst.FILE_UNKNOWN;
        if (proofNameUpper.indexOf('.PDF') !== -1) {
            xmpFileType = XMPConst.FILE_PDF;
        } else if (proofNameUpper.indexOf('.AI') !== -1) {
            xmpFileType = XMPConst.FILE_ILLUSTRATOR;
        } else if (proofNameUpper.indexOf('.EPS') !== -1) {
            xmpFileType = XMPConst.FILE_EPS;
        }

        var xmpFile = new XMPFile(proofFile.fsName, xmpFileType, XMPConst.OPEN_FOR_READ);
        if (!xmpFile) {
            return { data: null, status: 'Could not open XMP for: ' + proofFile.name };
        }

        var xmp = xmpFile.getXMP();
        var packingProp = xmp.getProperty('http://extremecolor.net/packing/', 'data');

        xmpFile.closeFile(XMPConst.CLOSE_UPDATE_IF_SAFE);

        if (packingProp && packingProp.value) {
            try {
                var parsed = JSON.parse(packingProp.value);
                return { data: parsed, status: 'Loaded from: ' + proofFile.name };
            } catch (parseError) {
                return { data: null, status: 'Found XMP in ' + proofFile.name + ' but JSON parse failed: ' + parseError.toString() };
            }
        }

        return { data: null, status: 'Proof found (' + proofFile.name + ') but no packing data in XMP' };

    } catch (e) {
        return { data: null, status: 'Error: ' + e.toString() };
    }
}

function buildManifestComparison(packingData, individualResults) {
    // Returns structured data for building the UI, not a text string
    var result = {
        hasData: false,
        expectedText: '',
        foundText: '',
        mismatches: []
    };

    if (!packingData) {
        return result;
    }
    result.hasData = true;

    // Build expected sizes by material from packing data
    var expectedSizes = {};
    if (packingData.materials) {
        for (var m = 0; m < packingData.materials.length; m++) {
            var material = packingData.materials[m];
            var materialName = material.material_name || 'Unknown';
            if (!expectedSizes[materialName]) { expectedSizes[materialName] = {}; }
            if (material.sheets) {
                for (var s = 0; s < material.sheets.length; s++) {
                    var sheet = material.sheets[s];
                    if (sheet.signs) {
                        for (var sg = 0; sg < sheet.signs.length; sg++) {
                            var sign = sheet.signs[sg];
                            var signW = sign.width || 0;
                            var signH = sign.height || 0;
                            var size = normalizeDimension(signW, signH);
                            if (signW > 0 && signH > 0) {
                                expectedSizes[materialName][size] = (expectedSizes[materialName][size] || 0) + 1;
                            }
                        }
                    }
                }
            }
        }
    }

    // Build found sizes by material from file check results
    // Also track which materials use CutOnly/PnC (uncountable via spot color)
    var foundSizes = {};
    var uncountableMaterials = {};
    if (individualResults) {
        for (var d = 0; d < individualResults.length; d++) {
            var ir = individualResults[d];
            var fileMat = ir.materialName || '';

            // Does this result carry per-artboard material buckets? (mixed PRIME files)
            var byMat = ir.cutThroughSizesByMaterial;
            var hasByMat = false;
            if (byMat) { for (var bmk in byMat) { if (byMat.hasOwnProperty(bmk)) { hasByMat = true; break; } } }

            if (fileMat.length > 0) {
                // Single material known for the whole file (split files, or a PRIME
                // file whose artboards all share one material). Unchanged behavior.
                if (!foundSizes[fileMat]) { foundSizes[fileMat] = {}; }
                if (ir.fileType === 'CutOnly' || ir.fileType === 'PnC') {
                    uncountableMaterials[fileMat] = true;
                }
                if (ir.cutThroughSizes) {
                    for (var size in ir.cutThroughSizes) {
                        foundSizes[fileMat][size] = (foundSizes[fileMat][size] || 0) + ir.cutThroughSizes[size];
                    }
                }
            } else if (hasByMat) {
                // Mixed-material PRIME file: file name gave no material, so bucket
                // each cut path under the material of the artboard it sits on.
                for (var pm in byMat) {
                    if (!byMat.hasOwnProperty(pm)) { continue; }
                    if (!foundSizes[pm]) { foundSizes[pm] = {}; }
                    for (var psize in byMat[pm]) {
                        foundSizes[pm][psize] = (foundSizes[pm][psize] || 0) + byMat[pm][psize];
                    }
                }
            } else {
                // No material info at all -> fall back to the old Unknown bucket.
                if (!foundSizes['Unknown']) { foundSizes['Unknown'] = {}; }
                if (ir.cutThroughSizes) {
                    for (var usize in ir.cutThroughSizes) {
                        foundSizes['Unknown'][usize] = (foundSizes['Unknown'][usize] || 0) + ir.cutThroughSizes[usize];
                    }
                }
            }
        }
    }

    // Collect all materials
    var allMaterials = {};
    var mat;
    for (mat in expectedSizes) { allMaterials[mat] = true; }
    for (mat in foundSizes) { allMaterials[mat] = true; }
    var materialList = [];
    for (mat in allMaterials) { materialList.push(mat); }
    materialList.sort();

    var expectedLines = [];
    var foundLines = [];

    for (var m = 0; m < materialList.length; m++) {
        var matName = materialList[m];
        var expectedByMat = expectedSizes[matName] || {};
        var foundByMat = foundSizes[matName] || {};

        // Material header
        expectedLines.push(matName);
        foundLines.push(matName);

        // If this material uses CutOnly/PnC, skip counting -- show note instead
        if (uncountableMaterials[matName]) {
            var expCount = 0;
            for (var sz in expectedByMat) { expCount += expectedByMat[sz]; }
            if (expCount > 0) {
                expectedLines.push('  ' + expCount + ' sign(s) expected');
            }
            foundLines.push('  (CutContour - not counted)');
            expectedLines.push('');
            foundLines.push('');
            continue;
        }

        // Collect all sizes for this material
        var allSizes = {};
        var sz;
        for (sz in expectedByMat) { allSizes[sz] = true; }
        for (sz in foundByMat) { allSizes[sz] = true; }
        var sizeList = [];
        for (sz in allSizes) { sizeList.push(sz); }
        sizeList.sort();

        // Matched items first
        for (var s = 0; s < sizeList.length; s++) {
            var size = sizeList[s];
            var expCount = expectedByMat[size] || 0;
            var fndCount = foundByMat[size] || 0;
            if (expCount > 0 && fndCount > 0 && expCount === fndCount) {
                expectedLines.push('  ' + expCount + 'x  ' + size);
                foundLines.push('  ' + fndCount + 'x  ' + size);
            }
        }

        // Mismatched items second
        for (var s = 0; s < sizeList.length; s++) {
            var size = sizeList[s];
            var expCount = expectedByMat[size] || 0;
            var fndCount = foundByMat[size] || 0;
            if (expCount === fndCount && expCount > 0) { continue; }

            if (expCount > 0 && fndCount > 0 && expCount !== fndCount) {
                expectedLines.push('  ' + expCount + 'x  ' + size + '  <<<');
                foundLines.push('  ' + fndCount + 'x  ' + size + '  <<<');
                result.mismatches.push(matName + ' ' + size + ': expected ' + expCount + ', found ' + fndCount);
            } else if (expCount > 0 && fndCount === 0) {
                expectedLines.push('  ' + expCount + 'x  ' + size + '  <<<');
                foundLines.push('  --missing--');
                result.mismatches.push('MISSING: ' + expCount + 'x ' + size + ' ' + matName);
            } else if (expCount === 0 && fndCount > 0) {
                expectedLines.push('  --not expected--');
                foundLines.push('  ' + fndCount + 'x  ' + size + '  <<<');
                result.mismatches.push('UNEXPECTED: ' + fndCount + 'x ' + size + ' ' + matName);
            }
        }

        // Blank line between materials
        expectedLines.push('');
        foundLines.push('');
    }

    result.expectedText = expectedLines.join('\n');
    result.foundText = foundLines.join('\n');
    return result;
}

function runAnalysis(documents, includePPI) {
    try {
        var overallStartTime = new Date().getTime();
        var timingReport = [];
        var totalImageCount = 0;
        var allLowResImages = [];
        var allImageDetails = [];
        var totalCutThroughPaths = 0;
        var totalCompoundPaths = 0;
        var allCutThroughSizes = {};
        var allCompoundPathWarnings = [];
        var documentNames = [];
        var individualDocumentResults = [];

        // Get the scale factor for Large Canvas handling
        var scaleFactor = 1;
        try {
            scaleFactor = app.activeDocument.scaleFactor || 1;
        } catch (e) {
            scaleFactor = 1;
        }

        // Store original active document
        var originalActiveDoc = app.activeDocument;

        // Try to read proof manifest
        var manifestResult = findAndReadProofManifest(documents);
        var proofManifest = manifestResult.data;
        var manifestStatus = manifestResult.status;

        // Process each document
        for (var docIndex = 0; docIndex < documents.length; docIndex++) {
            var doc = documents[docIndex];
            documentNames.push(doc.name);

            var docStartTime = new Date().getTime();

            // Activate document for analysis
            var documentActivated = false;
            var targetDoc = null;

            for (var activateIndex = 0; activateIndex < app.documents.length; activateIndex++) {
                if (app.documents[activateIndex].name === doc.name) {
                    targetDoc = app.documents[activateIndex];
                    app.activeDocument = targetDoc;
                    documentActivated = true;
                    break;
                }
            }

            if (!documentActivated) {
                alert("ERROR: Could not activate document: " + doc.name);
                continue;
            }

            // Force redraw to ensure document is active
            app.redraw();
            $.sleep(100);

            var docResult = analyzeDocument(targetDoc, includePPI, scaleFactor);
            var analysisTime = new Date().getTime() - docStartTime;

            // Determine material name: filename token first, then fall back to
            // artboard names (handles combined/PRIME files with no _PRINT_/_CUT_ token)
            var detectedMaterial = extractMaterialFromFilename(doc.name);
            if (!detectedMaterial || detectedMaterial.length === 0) {
                detectedMaterial = extractMaterialFromArtboards(targetDoc);
            }

            // Store individual document results
            var individualResult = {
                name: doc.name,
                materialName: detectedMaterial,
                fileType: extractFileType(doc.name),
                analysisTime: analysisTime,
                rasterCount: docResult.rasterCount,
                lowResImages: docResult.lowResImages || [],
                allImageDetails: docResult.allImageDetails || [],
                cutThroughSizes: docResult.cutThroughSizes || {},
                cutThroughSizesByMaterial: docResult.cutThroughSizesByMaterial || {},
                totalCutThroughPaths: docResult.totalCutThroughPaths || 0,
                compoundPathWarnings: docResult.compoundPathWarnings || [],
                totalCompoundPaths: docResult.totalCompoundPaths || 0,
                scaleFactor: scaleFactor
            };
            individualDocumentResults.push(individualResult);

            // Aggregate results
            totalImageCount += docResult.rasterCount;
            totalCutThroughPaths += docResult.totalCutThroughPaths;
            totalCompoundPaths += docResult.totalCompoundPaths;

            // Merge cut through sizes
            for (var size in docResult.cutThroughSizes) {
                if (allCutThroughSizes[size]) {
                    allCutThroughSizes[size] += docResult.cutThroughSizes[size];
                } else {
                    allCutThroughSizes[size] = docResult.cutThroughSizes[size];
                }
            }

            // Merge compound path warnings
            if (docResult.compoundPathWarnings) {
                for (var i = 0; i < docResult.compoundPathWarnings.length; i++) {
                    allCompoundPathWarnings.push({
                        document: doc.name,
                        text: docResult.compoundPathWarnings[i].text
                    });
                }
            }

            // Merge low res images
            if (docResult.lowResImages) {
                for (var i = 0; i < docResult.lowResImages.length; i++) {
                    allLowResImages.push({
                        document: doc.name,
                        text: docResult.lowResImages[i].text
                    });
                }
            }

            // Merge all image details
            if (docResult.allImageDetails) {
                for (var i = 0; i < docResult.allImageDetails.length; i++) {
                    allImageDetails.push({
                        document: doc.name,
                        text: docResult.allImageDetails[i].text
                    });
                }
            }

            timingReport.push("Document '" + doc.name + "': " + analysisTime + "ms");

            // Cleanup between documents
            if (docIndex < documents.length - 1) {
                try {
                    for (var cleanupIndex = 0; cleanupIndex < app.documents.length; cleanupIndex++) {
                        try {
                            app.documents[cleanupIndex].selection = null;
                        } catch (e) {}
                    }
                    app.redraw();
                    $.gc();
                    $.sleep(200);
                } catch (cleanupError) {}
            }
        }

        // Restore original active document
        try {
            app.activeDocument = originalActiveDoc;
        } catch (e) {}

        var totalTime = new Date().getTime() - overallStartTime;

        // Build report text
        var reportText = buildReport(documents.length, documentNames, totalTime, timingReport,
                                   totalImageCount, allLowResImages, allImageDetails, allCutThroughSizes,
                                   totalCutThroughPaths, totalCompoundPaths, allCompoundPathWarnings,
                                   includePPI, individualDocumentResults);

        // Build manifest comparison if available
        var manifestComparison = buildManifestComparison(proofManifest, individualDocumentResults);

        // Store analysis data
        var analysisData = {
            reportText: reportText,
            manifestComparison: manifestComparison,
            individualResults: individualDocumentResults,
            totalDocuments: documents.length,
            totalPaths: totalCutThroughPaths,
            totalCompoundPaths: totalCompoundPaths,
            totalTime: Math.round(totalTime / 1000),
            hasLowRes: includePPI && allLowResImages.length > 0,
            hasCompoundPaths: totalCompoundPaths > 0,
            manifestStatus: manifestStatus,
            allCutThroughSizes: allCutThroughSizes,
            allImageDetails: allImageDetails,
            allCompoundPathWarnings: allCompoundPathWarnings
        };

        // Show results
        showResultsDialog(analysisData);

    } catch (error) {
        alert("Error in analysis: " + error.toString());
    }
}

function analyzeDocument(doc, includePPI, scaleFactor) {
    function pointsToInches(points) {
        try {
            var docScaleFactor = doc.scaleFactor;
            if (!docScaleFactor || docScaleFactor <= 0 || isNaN(docScaleFactor)) {
                docScaleFactor = 1;
            }
            return Math.round(((points / 72) * docScaleFactor) * 100) / 100;
        } catch (e) {
            return Math.round((points / 72) * 100) / 100;
        }
    }

    function getPathDimensions(pathItem) {
        var bounds = pathItem.geometricBounds;
        var width = pointsToInches(bounds[2] - bounds[0]);
        var height = pointsToInches(bounds[1] - bounds[3]);

        // Round to nearest 1/8" (0.125)
        width = Math.round(width * 8) / 8;
        height = Math.round(height * 8) / 8;

        if (width <= height) {
            return width + '"x' + height + '"';
        } else {
            return height + '"x' + width + '"';
        }
    }

    function findPathsWithCutThroughColor(doc) {
        var cutThroughSizes = {};
        // Cut sizes grouped by the material of the ARTBOARD each path sits on.
        // This is what lets mixed-material PRIME files match the proof (each board
        // is one material, so we look up the board a path is on and use its name).
        var cutThroughSizesByMaterial = {};
        var totalPaths = 0;
        var compoundPathWarnings = [];
        var totalCompoundPaths = 0;

        // Precompute each artboard's rectangle + material name (name minus _Pt#/_Part#).
        var abInfo = [];
        try {
            for (var ab = 0; ab < doc.artboards.length; ab++) {
                var abName = doc.artboards[ab].name || '';
                var abMat = abName.replace(/_Pt\d+$/i, '').replace(/_Part\d+$/i, '').replace(/^\s+|\s+$/g, '');
                if (abMat.length === 0) { abMat = 'Unknown'; }
                abInfo.push({ rect: doc.artboards[ab].artboardRect, material: abMat });
            }
        } catch (eAB) {}

        // Which artboard (material) contains a path, by its center point.
        // artboardRect and geometricBounds are both [L, T, R, B] in the same
        // point space, so scaleFactor is irrelevant to this containment test.
        function materialForBounds(b) {
            var cx = (b[0] + b[2]) / 2;
            var cy = (b[1] + b[3]) / 2;
            for (var q = 0; q < abInfo.length; q++) {
                var r = abInfo[q].rect;
                if (cx >= r[0] && cx <= r[2] && cy <= r[1] && cy >= r[3]) {
                    return abInfo[q].material;
                }
            }
            return 'Unknown';
        }

        try {
            // Verify the spot color exists in this document
            var hasSpot = false;
            for (var i = 0; i < doc.spots.length; i++) {
                if (doc.spots[i].name == "CutThrough2-Outside") {
                    hasSpot = true;
                    break;
                }
            }

            if (!hasSpot) {
                return {
                    cutThroughSizes: cutThroughSizes,
                    cutThroughSizesByMaterial: cutThroughSizesByMaterial,
                    totalCutThroughPaths: 0,
                    compoundPathWarnings: compoundPathWarnings,
                    totalCompoundPaths: 0
                };
            }

            // Check if a color references the CutThrough2-Outside spot
            function isCutThroughColor(color) {
                try {
                    if (color && color.typename == "SpotColor" && color.spot && color.spot.name == "CutThrough2-Outside") {
                        return true;
                    }
                } catch (e) {}
                return false;
            }

            // Check if a path item has CutThrough color on fill or stroke
            function pathHasCutThrough(pathItem) {
                try {
                    if (pathItem.filled && isCutThroughColor(pathItem.fillColor)) return true;
                } catch (e) {}
                try {
                    if (pathItem.stroked && isCutThroughColor(pathItem.strokeColor)) return true;
                } catch (e) {}
                return false;
            }

            var allPaths = [];

            // Recursive traversal of all page items
            function walkItems(container) {
                var items;
                try {
                    items = container.pageItems;
                } catch (e) {
                    return;
                }

                for (var i = 0; i < items.length; i++) {
                    var item = items[i];
                    try {
                        if (item.typename == "PathItem") {
                            if (pathHasCutThrough(item)) {
                                allPaths.push(item);
                            }
                        } else if (item.typename == "CompoundPathItem") {
                            // Check sub-paths for the spot color
                            var hasColor = false;
                            try {
                                for (var p = 0; p < item.pathItems.length; p++) {
                                    if (pathHasCutThrough(item.pathItems[p])) {
                                        hasColor = true;
                                        break;
                                    }
                                }
                            } catch (e) {}
                            if (hasColor) {
                                totalCompoundPaths++;
                                var size = getPathDimensions(item);
                                compoundPathWarnings.push({
                                    text: "WARNING: CutThrough color on COMPOUND PATH (size: " + size + ") - This will cause production errors!"
                                });
                            }
                        } else if (item.typename == "GroupItem") {
                            walkItems(item);
                        }
                    } catch (e) {}
                }
            }

            // Walk all layers (including locked/hidden -- we still want to count them)
            for (var layerIdx = 0; layerIdx < doc.layers.length; layerIdx++) {
                walkItems(doc.layers[layerIdx]);
            }

            // Tally sizes (overall and per artboard-material)
            for (var i = 0; i < allPaths.length; i++) {
                var size = getPathDimensions(allPaths[i]);
                if (cutThroughSizes[size]) {
                    cutThroughSizes[size]++;
                } else {
                    cutThroughSizes[size] = 1;
                }
                var mat = materialForBounds(allPaths[i].geometricBounds);
                if (!cutThroughSizesByMaterial[mat]) { cutThroughSizesByMaterial[mat] = {}; }
                cutThroughSizesByMaterial[mat][size] = (cutThroughSizesByMaterial[mat][size] || 0) + 1;
                totalPaths++;
            }

        } catch (e) {}

        return {
            cutThroughSizes: cutThroughSizes,
            cutThroughSizesByMaterial: cutThroughSizesByMaterial,
            totalCutThroughPaths: totalPaths,
            compoundPathWarnings: compoundPathWarnings,
            totalCompoundPaths: totalCompoundPaths
        };
    }

    // Image analysis
    var rasterCount = 0;
    var lowResImages = [];
    var allImageDetails = [];

    // FAST PRE-CHECK: Skip image analysis if document has no images
    var hasImages = false;
    try {
        if (doc.rasterItems.length > 0 || doc.placedItems.length > 0) {
            hasImages = true;
        }
    } catch (e) {
        hasImages = true;
    }

    if (includePPI && hasImages) {
        try {
            var processedItems = [];

            function processImageItem(item) {
                if (item.typename == "RasterItem" || item.typename == "PlacedItem") {
                    var alreadyProcessed = false;
                    for (var k = 0; k < processedItems.length; k++) {
                        if (processedItems[k] === item) {
                            alreadyProcessed = true;
                            break;
                        }
                    }

                    if (!alreadyProcessed) {
                        processedItems.push(item);
                        rasterCount++;

                        try {
                            // Get basic image info with Large Canvas correction
                            var bounds = item.geometricBounds;
                            var actualScaleFactor = (scaleFactor && scaleFactor > 0) ? scaleFactor : 1;
                            var docWidthInches = Math.round(((bounds[2] - bounds[0]) / 72 * scaleFactor) * 100) / 100;
                            var docHeightInches = Math.round(((bounds[1] - bounds[3]) / 72 * scaleFactor) * 100) / 100;

                            var imageInfo = {
                                type: item.typename,
                                bounds: bounds,
                                docWidth: docWidthInches,
                                docHeight: docHeightInches
                            };

                            // Try multiple methods to get PPI/resolution info
                            var estimatedPPI = "Unknown";
                            var resolutionMethod = "No method available";

                            // Method 1: Try to get actual resolution for RasterItems
                            if (item.typename == "RasterItem") {
                                try {
                                    var fileStatus = "Embedded";
                                    try {
                                        if (item.file && item.file.exists) {
                                            fileStatus = "Linked: " + item.file.name;
                                        }
                                    } catch (fileError) {
                                        fileStatus = "Embedded";
                                    }

                                    resolutionMethod = fileStatus;

                                    // Calculate from transformation matrix
                                    try {
                                        var matrix = item.matrix;
                                        if (matrix && matrix.mValueA !== undefined && matrix.mValueD !== undefined) {
                                            var scaleX = Math.abs(matrix.mValueA);
                                            var scaleY = Math.abs(matrix.mValueD);
                                            var avgScale = (scaleX + scaleY) / 2;
                                            if (avgScale > 0) {
                                                var basePPI = 72 / avgScale;
                                                var correctedPPI = Math.round(basePPI / scaleFactor);
                                                estimatedPPI = correctedPPI;
                                                resolutionMethod += " (Matrix: scaleX=" + scaleX.toFixed(3) + ", scaleY=" + scaleY.toFixed(3) + ", avg=" + avgScale.toFixed(3);
                                                if (actualScaleFactor !== 1) {
                                                    resolutionMethod += ", Large Canvas factor=" + actualScaleFactor + ", corrected PPI=" + correctedPPI;
                                                }
                                                resolutionMethod += ")";
                                            } else {
                                                resolutionMethod += " (Matrix: invalid scale values)";
                                            }
                                        } else {
                                            resolutionMethod += " (Matrix: no matrix data)";
                                        }
                                    } catch (matrixError) {
                                        resolutionMethod += " (Matrix error: " + matrixError.toString() + ")";
                                    }
                                } catch (rasterError) {
                                    resolutionMethod = "RasterItem analysis error: " + rasterError.toString();
                                }
                            }

                            // Method 2: Try for PlacedItems
                            else if (item.typename == "PlacedItem") {
                                try {
                                    var fileStatus = "Unknown";
                                    try {
                                        if (item.file && item.file.exists) {
                                            fileStatus = "Linked: " + item.file.name;
                                        } else {
                                            fileStatus = "Missing link";
                                        }
                                    } catch (fileError) {
                                        fileStatus = "No file reference";
                                    }

                                    resolutionMethod = fileStatus;

                                    // Calculate from transformation matrix
                                    try {
                                        var matrix = item.matrix;
                                        if (matrix && matrix.mValueA !== undefined && matrix.mValueD !== undefined) {
                                            var scaleX = Math.abs(matrix.mValueA);
                                            var scaleY = Math.abs(matrix.mValueD);
                                            var avgScale = (scaleX + scaleY) / 2;
                                            if (avgScale > 0) {
                                                estimatedPPI = Math.round(72 / avgScale);
                                                resolutionMethod += " (Matrix: scaleX=" + scaleX.toFixed(3) + ", scaleY=" + scaleY.toFixed(3) + ", avg=" + avgScale.toFixed(3) + ")";
                                            } else {
                                                resolutionMethod += " (Matrix: invalid scale values)";
                                            }
                                        } else {
                                            resolutionMethod += " (Matrix: no matrix data)";
                                        }
                                    } catch (matrixError) {
                                        resolutionMethod += " (Matrix error: " + matrixError.toString() + ")";
                                    }
                                } catch (placedError) {
                                    resolutionMethod = "PlacedItem analysis error: " + placedError.toString();
                                }
                            }

                            // Build comprehensive image text
                            var imageText = 'Img#' + rasterCount + ' (' + item.typename + '): SIZE: ' + docWidthInches + '"x' + docHeightInches + '"';

                            if (estimatedPPI !== "Unknown") {
                                imageText += ', PPI: ~' + estimatedPPI;
                            } else {
                                imageText += ', PPI: Unknown';
                            }

                            imageText += ' [' + resolutionMethod + ']';

                            // Add to ALL images list
                            allImageDetails.push({text: imageText});

                            // Check if low resolution
                            var isLowRes = false;
                            if (estimatedPPI !== "Unknown" && typeof estimatedPPI === "number") {
                                isLowRes = (estimatedPPI < 72);
                                if (isLowRes) {
                                    imageText += ' LOW RESOLUTION!';
                                    lowResImages.push({text: imageText});
                                }
                            }

                        } catch (imageError) {
                            var errorText = 'Image ' + rasterCount + ' (' + item.typename + '): Analysis error - ' + imageError.toString();
                            allImageDetails.push({text: errorText});
                        }
                    }
                }

                if (item.typename == "GroupItem" && item.pageItems.length > 0) {
                    for (var j = 0; j < item.pageItems.length; j++) {
                        processImageItem(item.pageItems[j]);
                    }
                }
            }

            for (var i = 0; i < doc.pageItems.length; i++) {
                processImageItem(doc.pageItems[i]);
            }
        } catch (e) {
            allImageDetails.push({text: "PPI Analysis failed: " + e.toString()});
        }
    } else if (hasImages) {
        // Quick analysis - just count images
        try {
            rasterCount = doc.rasterItems.length + doc.placedItems.length;
            if (rasterCount > 0) {
                allImageDetails.push({text: "Quick scan found " + rasterCount + " images (no details in quick mode)"});
            }
        } catch (e) {
            allImageDetails.push({text: "Image counting failed: " + e.toString()});
        }
    }

    // CutThrough analysis - FAST METHOD using Find Color
    var cutThroughResults = findPathsWithCutThroughColor(doc);
    var cutThroughSizes = cutThroughResults.cutThroughSizes;
    var cutThroughSizesByMaterial = cutThroughResults.cutThroughSizesByMaterial || {};
    var totalCutThroughPaths = cutThroughResults.totalCutThroughPaths;
    var compoundPathWarnings = cutThroughResults.compoundPathWarnings;
    var totalCompoundPaths = cutThroughResults.totalCompoundPaths;

    return {
        rasterCount: rasterCount,
        lowResImages: lowResImages,
        allImageDetails: allImageDetails,
        cutThroughSizes: cutThroughSizes,
        cutThroughSizesByMaterial: cutThroughSizesByMaterial,
        totalCutThroughPaths: totalCutThroughPaths,
        compoundPathWarnings: compoundPathWarnings,
        totalCompoundPaths: totalCompoundPaths
    };
}

function showResultsDialog(analysisData) {
    var dialog = new Window("dialog", "Analysis Results");
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.spacing = 8;
    dialog.margins = 12;
    dialog.preferredSize.width = 1100;

    // -- ROW 1: Summary header --
    var headerGroup = dialog.add("group");
    headerGroup.orientation = "row";
    headerGroup.alignment = "fill";
    headerGroup.spacing = 20;
    headerGroup.add("statictext", undefined, "Documents: " + analysisData.totalDocuments);
    headerGroup.add("statictext", undefined, "Cut Paths: " + analysisData.totalPaths);
    headerGroup.add("statictext", undefined, "Time: " + analysisData.totalTime + "s");

    // Critical warnings inline
    if (analysisData.hasCompoundPaths) {
        var cpWarn = headerGroup.add("statictext", undefined, "COMPOUND PATHS: " + analysisData.totalCompoundPaths + " - FIX!");
        cpWarn.graphics.foregroundColor = cpWarn.graphics.newPen(cpWarn.graphics.PenType.SOLID_COLOR, [1, 0, 0], 1);
        cpWarn.graphics.font = ScriptUI.newFont("dialog", "Bold", 11);
    }
    if (analysisData.hasLowRes) {
        var lrCount = 0;
        for (var lr = 0; lr < analysisData.individualResults.length; lr++) {
            if (analysisData.individualResults[lr].lowResImages) {
                lrCount += analysisData.individualResults[lr].lowResImages.length;
            }
        }
        var lrWarn = headerGroup.add("statictext", undefined, "LOW RES: " + lrCount);
        lrWarn.graphics.foregroundColor = lrWarn.graphics.newPen(lrWarn.graphics.PenType.SOLID_COLOR, [1, 0, 0], 1);
        lrWarn.graphics.font = ScriptUI.newFont("dialog", "Bold", 11);
    }

    // -- ROW 2: Manifest comparison (always visible) --
    var comparison = analysisData.manifestComparison;

    if (comparison && comparison.hasData) {
        // Mismatch alerts
        if (comparison.mismatches.length > 0) {
            var alertPanel = dialog.add("panel", undefined, "MISMATCHES");
            alertPanel.orientation = "column";
            alertPanel.alignChildren = "left";
            alertPanel.margins = [10, 15, 10, 8];
            for (var mi = 0; mi < comparison.mismatches.length; mi++) {
                var alertLine = alertPanel.add("statictext", undefined, comparison.mismatches[mi]);
                alertLine.graphics.foregroundColor = alertLine.graphics.newPen(alertLine.graphics.PenType.SOLID_COLOR, [1, 0.15, 0.15], 1);
                alertLine.graphics.font = ScriptUI.newFont("dialog", "Bold", 11);
            }
        } else {
            var okGroup = dialog.add("group");
            okGroup.alignment = "center";
            var okText = okGroup.add("statictext", undefined, "All expected items found - quantities match");
            okText.graphics.foregroundColor = okText.graphics.newPen(okText.graphics.PenType.SOLID_COLOR, [0, 0.55, 0], 1);
            okText.graphics.font = ScriptUI.newFont("dialog", "Bold", 12);
        }

        // Side by side panels
        var compGroup = dialog.add("group");
        compGroup.orientation = "row";
        compGroup.alignChildren = "fill";
        compGroup.spacing = 10;

        var expectedPanel = compGroup.add("panel", undefined, "EXPECTED (from proof)");
        expectedPanel.orientation = "column";
        expectedPanel.alignChildren = "fill";
        expectedPanel.preferredSize.width = 520;
        expectedPanel.margins = [10, 15, 10, 8];
        var expectedEdit = expectedPanel.add("edittext", undefined, comparison.expectedText, {multiline: true, scrolling: true, readonly: true});
        expectedEdit.preferredSize.height = 180;

        var foundPanel = compGroup.add("panel", undefined, "FOUND (in files)");
        foundPanel.orientation = "column";
        foundPanel.alignChildren = "fill";
        foundPanel.preferredSize.width = 520;
        foundPanel.margins = [10, 15, 10, 8];
        var foundEdit = foundPanel.add("edittext", undefined, comparison.foundText, {multiline: true, scrolling: true, readonly: true});
        foundEdit.preferredSize.height = 180;

    } else {
        // No manifest data - show compact status
        var noDataGroup = dialog.add("group");
        noDataGroup.alignment = "left";
        var noDataText = noDataGroup.add("statictext", undefined, "Manifest: " + (analysisData.manifestStatus || 'No proof data found'));
        noDataText.graphics.font = ScriptUI.newFont("dialog", "Regular", 10);
    }

    // -- ROW 3: Document details (sidebar + content) --
    var contentGroup = dialog.add("group");
    contentGroup.orientation = "row";
    contentGroup.alignChildren = "fill";

    // LEFT SIDEBAR
    var sidebarGroup = contentGroup.add("group");
    sidebarGroup.orientation = "column";
    sidebarGroup.alignChildren = "fill";
    sidebarGroup.preferredSize.width = 220;

    var sidebarLabel = sidebarGroup.add("statictext", undefined, "Documents:");
    sidebarLabel.graphics.font = ScriptUI.newFont("dialog", "Bold", 11);

    var listBox = sidebarGroup.add("listbox");
    listBox.alignment = ["fill", "fill"];

    // Build listbox items: Full Report + individual documents
    var listItems = [];
    listItems.push("Full Report");

    if (analysisData.individualResults) {
        for (var d = 0; d < analysisData.individualResults.length; d++) {
            var result = analysisData.individualResults[d];
            var displayName = result.materialName && result.materialName.length > 0 ? result.materialName : result.name;
            listItems.push(displayName);
        }
    }

    for (var li = 0; li < listItems.length; li++) {
        listBox.add("item", listItems[li]);
    }
    listBox.selection = 0;

    // RIGHT CONTENT
    var contentDisplayGroup = contentGroup.add("group");
    contentDisplayGroup.orientation = "column";
    contentDisplayGroup.alignChildren = "fill";
    contentDisplayGroup.preferredSize.width = 830;

    var contentText = contentDisplayGroup.add("edittext", undefined, "", {multiline: true, scrolling: true, readonly: true});
    contentText.alignment = ["fill", "fill"];
    contentText.preferredSize.height = 300;

    // Build content panels
    var contentPanels = [];

    // Panel 0: Full Report
    contentPanels.push(analysisData.reportText);

    // Panel 1+: Individual documents
    if (analysisData.individualResults) {
        for (var d = 0; d < analysisData.individualResults.length; d++) {
            var docResult = analysisData.individualResults[d];
            var docReport = [];
            docReport.push("DOCUMENT: " + docResult.name);
            docReport.push("Material: " + (docResult.materialName && docResult.materialName.length > 0 ? docResult.materialName : "Unknown"));
            docReport.push("Processing: " + docResult.analysisTime + "ms");
            docReport.push("Images: " + docResult.rasterCount);
            docReport.push("Cut Paths: " + docResult.totalCutThroughPaths);

            if (docResult.compoundPathWarnings && docResult.compoundPathWarnings.length > 0) {
                docReport.push("");
                docReport.push("*** CRITICAL: COMPOUND PATH ERRORS ***");
                for (var w = 0; w < docResult.compoundPathWarnings.length; w++) {
                    docReport.push(docResult.compoundPathWarnings[w].text);
                }
            }

            docReport.push("");

            if (docResult.totalCutThroughPaths > 0) {
                docReport.push("CUTTHROUGH BREAKDOWN");
                var docSizes = [];
                for (var size in docResult.cutThroughSizes) {
                    docSizes.push(size);
                }
                docSizes.sort();
                for (var s = 0; s < docSizes.length; s++) {
                    var size = docSizes[s];
                    var count = docResult.cutThroughSizes[size];
                    docReport.push("  " + count + "x  " + size);
                }
            } else {
                docReport.push("No CutThrough paths found");
            }

            if (docResult.allImageDetails && docResult.allImageDetails.length > 0) {
                docReport.push("");
                docReport.push("ALL IMAGE DETAILS");
                for (var img = 0; img < docResult.allImageDetails.length; img++) {
                    docReport.push(docResult.allImageDetails[img].text);
                }
            }

            if (docResult.lowResImages && docResult.lowResImages.length > 0) {
                docReport.push("");
                docReport.push("*** LOW RESOLUTION IMAGES ***");
                for (var img = 0; img < docResult.lowResImages.length; img++) {
                    docReport.push(docResult.lowResImages[img].text);
                }
            }

            contentPanels.push(docReport.join("\n"));
        }
    }

    // Listbox change handler
    listBox.onChange = function() {
        var selectedIndex = listBox.selection ? listBox.selection.index : 0;
        if (selectedIndex >= 0 && selectedIndex < contentPanels.length) {
            contentText.text = contentPanels[selectedIndex];
        }
    };

    // Initialize with first panel
    if (contentPanels.length > 0) {
        contentText.text = contentPanels[0];
    }

    // Buttons
    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "center";
    buttonGroup.spacing = 10;

    var exportBtn = buttonGroup.add("button", undefined, "Export Report");
    var closeBtn = buttonGroup.add("button", undefined, "Close");

    exportBtn.onClick = function() {
        try {
            var saveFile = File.saveDialog("Save report as text file", "Text files:*.txt");
            if (saveFile) {
                saveFile.open("w");
                saveFile.write(analysisData.reportText);
                saveFile.close();
                alert("Report saved successfully!");
            }
        } catch (e) {
            alert("Error saving file: " + e.toString());
        }
    };

    closeBtn.onClick = function() {
        dialog.close();
    };

    dialog.show();
}

function buildReport(docCount, documentNames, totalTime, timingReport,
                    totalImageCount, allLowResImages, allImageDetails, allCutThroughSizes,
                    totalCutThroughPaths, totalCompoundPaths, allCompoundPathWarnings,
                    includePPI, individualDocumentResults) {
    var report = [];
    report.push("=== MULTI-DOCUMENT ANALYSIS SUMMARY ===");
    report.push("");

    var hasLargeCanvas = false;
    for (var i = 0; i < individualDocumentResults.length; i++) {
        if (individualDocumentResults[i].scaleFactor && individualDocumentResults[i].scaleFactor !== 1) {
            hasLargeCanvas = true;
            break;
        }
    }

    if (hasLargeCanvas) {
        report.push("*** LARGE CANVAS DOCUMENTS DETECTED ***");
        report.push("Some documents use Large Canvas mode. Measurements have been corrected.");
        report.push("");
    }

    if (totalCompoundPaths > 0) {
        report.push("*** CRITICAL PRODUCTION ERROR ***");
        report.push("COMPOUND PATHS WITH CUTTHROUGH COLOR DETECTED!");
        report.push("Found " + totalCompoundPaths + " compound path(s) with CutThrough color");
        report.push("This WILL cause cutting machine failures!");
        report.push("MUST BE FIXED before production!");
        report.push("");
        report.push("Compound Path Details:");
        for (var i = 0; i < allCompoundPathWarnings.length; i++) {
            report.push("  " + allCompoundPathWarnings[i].document + ": " + allCompoundPathWarnings[i].text);
        }
        report.push("");
    }

    report.push("Analysis Type: " + (includePPI ? "Full Analysis (with PPI calculations)" : "Quick Analysis (no PPI calculations)"));
    report.push("Documents Analyzed: " + docCount);
    report.push("Total Processing Time: " + Math.round(totalTime / 1000) + " seconds");
    report.push("");

    report.push("Documents Processed:");
    for (var i = 0; i < documentNames.length; i++) {
        report.push("  " + (i + 1) + ". " + documentNames[i]);
    }
    report.push("");

    report.push("=== KEY METRICS ===");
    report.push("Total Images: " + totalImageCount);
    if (includePPI && allLowResImages.length > 0) {
        report.push("Low Resolution Images (< 72 PPI): " + allLowResImages.length + " ** WARNING **");
    } else if (includePPI) {
        report.push("Low Resolution Images (< 72 PPI): 0 (All images meet requirements)");
    }
    report.push("Total CutThrough2-Outside Paths: " + totalCutThroughPaths);
    if (totalCompoundPaths > 0) {
        report.push("Compound Paths with CutThrough Color: " + totalCompoundPaths + " ** CRITICAL ERROR **");
    }
    report.push("");

    report.push("=== PER-DOCUMENT BREAKDOWN ===");

    var sortedResults = [];
    for (var i = 0; i < individualDocumentResults.length; i++) {
        sortedResults.push(individualDocumentResults[i]);
    }
    sortedResults.sort(function(a, b) {
        var nameA = a.name.toUpperCase();
        var nameB = b.name.toUpperCase();
        if (nameA < nameB) return -1;
        if (nameA > nameB) return 1;
        return 0;
    });

    for (var i = 0; i < sortedResults.length; i++) {
        var docResult = sortedResults[i];
        report.push("");
        report.push("Document: " + docResult.name);
        if (docResult.scaleFactor && docResult.scaleFactor !== 1) {
            report.push("  ** Large Canvas (Scale Factor: " + docResult.scaleFactor + ") **");
        }

        if (docResult.compoundPathWarnings && docResult.compoundPathWarnings.length > 0) {
            report.push("  ** COMPOUND PATH ERRORS DETECTED **");
            for (var w = 0; w < docResult.compoundPathWarnings.length; w++) {
                report.push("    " + docResult.compoundPathWarnings[w].text);
            }
        }

        if (docResult.totalCutThroughPaths > 0) {
            var docSizes = [];
            for (var size in docResult.cutThroughSizes) {
                docSizes.push(size);
            }
            docSizes.sort();
            for (var s = 0; s < docSizes.length; s++) {
                var size = docSizes[s];
                var count = docResult.cutThroughSizes[size];
                report.push("    " + count + " x " + size);
            }
        } else {
            report.push("  No CutThrough paths found");
        }
    }
    report.push("");

    if (includePPI && allImageDetails && allImageDetails.length > 0) {
        report.push("=== ALL IMAGE DETAILS ===");
        for (var i = 0; i < allImageDetails.length; i++) {
            report.push(allImageDetails[i].text);
        }
        report.push("");
    }

    if (totalCutThroughPaths > 0) {
        report.push("=== CUTTHROUGH2-OUTSIDE BREAKDOWN ===");
        var sortedSizes = [];
        for (var size in allCutThroughSizes) {
            sortedSizes.push(size);
        }
        sortedSizes.sort();

        for (var i = 0; i < sortedSizes.length; i++) {
            var size = sortedSizes[i];
            var count = allCutThroughSizes[size];
            report.push(count + "\t" + size);
        }
        report.push("");
    }

    report.push("=== PERFORMANCE TIMING ===");
    report.push("Overall Processing: " + Math.round(totalTime) + "ms (" + Math.round(totalTime / 1000) + "s)");
    report.push("Average per Document: " + Math.round(totalTime / docCount) + "ms");
    report.push("");
    report.push("Individual Document Times:");
    for (var i = 0; i < timingReport.length; i++) {
        report.push("  " + timingReport[i]);
    }

    return report.join("\n");
}
