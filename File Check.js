#target illustrator

/*@METADATA{
  "name": "File Check",
  "description": "Production file readiness analysis with manifest comparison",
  "version": "4.3",
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
    // Base material from a production file name (see PROD NAME PARSER).
    // BlkBk, FRONT/BACK, parts and _x# are separate tokens, so
    // "..._PRINT_3mmACM_BlkBk_Pt8_x2.pdf" -> "3mmACM".
    return ProdName.parseFileName(filename).material;
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
            var base = ProdName.parseArtboardName(abName).material;
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
    return ProdName.parseFileName(filename).type;
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

function collectNearbyProductionNames(documents) {
    // File names of production files that sit near the analyzed documents, so a
    // CUT file can learn its PRINT file's _x# even when that PRINT isn't open.
    // Scans: each document's folder, its PRINT/CUT subfolders, and - when the
    // document lives in a PRINT or CUT subfolder - the parent and its siblings.
    var names = [];
    var seen = {};
    function addFolder(f) {
        try {
            if (!f || !f.exists) { return; }
            if (seen[f.fsName]) { return; }
            seen[f.fsName] = true;
            var files = f.getFiles();
            for (var i = 0; i < files.length; i++) {
                if (files[i] instanceof File) {
                    var nm = String(files[i].name);
                    try { nm = decodeURI(nm); } catch (eDec) {}
                    if (/^[._~]/.test(nm)) { continue; }
                    if (!/\.(ai|pdf|eps)$/i.test(nm)) { continue; }
                    names.push(nm);
                }
            }
        } catch (e) {}
    }
    function addProdSubfolders(f) {
        try {
            var subs = f.getFiles();
            for (var s = 0; s < subs.length; s++) {
                if (subs[s] instanceof Folder && /^(PRINT|CUT)$/i.test(String(subs[s].name))) {
                    addFolder(subs[s]);
                }
            }
        } catch (e) {}
    }
    for (var d = 0; d < documents.length; d++) {
        try {
            var folder = new File(documents[d].fullName).parent;
            if (!folder) { continue; }
            addFolder(folder);
            addProdSubfolders(folder);
            if (/^(PRINT|CUT)$/i.test(String(folder.name)) && folder.parent) {
                addFolder(folder.parent);
                addProdSubfolders(folder.parent);
            }
        } catch (e) {}
    }
    return names;
}

function isPrintType(t) {
    return t === 'PRINT' || t === 'PnC' || t === 'CutOnly';
}

function computeRunMultipliers(individualResults, nearbyNames) {
    // Sets on each result:
    //   runMult   - how many times this file's cut paths are produced
    //   countMode - 'normal' | 'back' (printed back, not counted) | 'skip'
    //   runNote   - human readable explanation for the report
    //   runWarnings - [] of strings
    var printRuns = {};     // pairKey -> times the PRINT sheet runs
    var printSrc = {};      // pairKey -> file name the run count came from
    var printHasCuts = {};  // pairKey -> PRINT file already holds the cut paths
    var i, j, p, parts, k;
    var frontAlias = {};    // no-side key -> FRONT print key (used only if no one-sided print has it)

    function partsOf(pp) { return pp.parts.length ? pp.parts : [0]; }

    for (i = 0; i < individualResults.length; i++) {
        p = individualResults[i].parsed;
        if (!p || !isPrintType(p.type)) { continue; }
        parts = partsOf(p);
        for (j = 0; j < parts.length; j++) {
            k = ProdName.pairKey(p, parts[j]);
            printRuns[k] = p.mult;
            printSrc[k] = individualResults[i].name;
            if (individualResults[i].totalCutThroughPaths > 0 && p.side !== 'BACK') {
                printHasCuts[k] = individualResults[i].name;
            }
            // 2-sided sheets: PRINT_..._FRONT_Pt14 pairs with CUT_..._Pt14
            if (p.side === 'FRONT') {
                var nk = ProdName.pairKeyNoSide(p, parts[j]);
                if (!frontAlias.hasOwnProperty(nk)) { frontAlias[nk] = k; }
            }
        }
    }
    for (i = 0; i < nearbyNames.length; i++) {
        p = ProdName.parseFileName(nearbyNames[i]);
        if (!isPrintType(p.type)) { continue; }
        parts = partsOf(p);
        for (j = 0; j < parts.length; j++) {
            k = ProdName.pairKey(p, parts[j]);
            if (!printRuns.hasOwnProperty(k)) {
                printRuns[k] = p.mult;
                printSrc[k] = nearbyNames[i];
            }
        }
    }

    for (i = 0; i < individualResults.length; i++) {
        var ir = individualResults[i];
        p = ir.parsed;
        ir.runMult = 1;
        ir.countMode = 'normal';
        ir.runNote = '';
        ir.runWarnings = [];
        if (!p) { continue; }

        if (isPrintType(p.type)) {
            if (p.side === 'BACK') {
                ir.countMode = 'back';
                ir.runMult = 0;
                ir.runNote = 'Printed back - checked for presence only, cuts not counted';
                if (ir.totalCutThroughPaths > 0) {
                    ir.runWarnings.push('BACK file has ' + ir.totalCutThroughPaths + ' cut path(s) - backs should not carry cuts');
                }
            } else {
                ir.runMult = p.mult;
                if (p.mult > 1) { ir.runNote = 'Sheet runs x' + p.mult + ' - cut counts multiplied'; }
            }
        } else if (p.type === 'CUT') {
            parts = partsOf(p);
            var total = 0;
            var used = [];
            var skipped = [];
            for (j = 0; j < parts.length; j++) {
                k = ProdName.pairKey(p, parts[j]);
                if (!p.side && !printRuns.hasOwnProperty(k) && frontAlias.hasOwnProperty(k)) { k = frontAlias[k]; }
                var ptLabel = parts[j] ? ('Pt' + parts[j]) : 'single part';
                if (printHasCuts[k]) {
                    skipped.push(ptLabel);
                    continue;
                }
                var runs = p.mult;
                if (printRuns.hasOwnProperty(k)) {
                    runs = printRuns[k];
                    if (p.hasMult && p.mult !== runs) {
                        ir.runWarnings.push('CUT says x' + p.mult + ' but ' + printSrc[k] + ' says x' + runs + ' - using x' + runs);
                    }
                }
                total += runs;
                used.push(ptLabel + ' x' + runs);
            }
            if (used.length === 0 && skipped.length > 0) {
                ir.countMode = 'skip';
                ir.runMult = 0;
                ir.runNote = 'Not counted - these cuts are already counted inside their PRINT file (' + skipped.join(', ') + ')';
            } else {
                ir.runMult = total;
                if (used.length > 1 || total !== 1) {
                    ir.runNote = 'Cut runs: ' + used.join(' + ') + ' = x' + total;
                }
                if (skipped.length > 0) {
                    ir.runNote += (ir.runNote ? '; ' : '') + 'already counted in PRINT: ' + skipped.join(', ');
                }
            }
        }
    }
}

function parseSizeKey(sizeKey) {
    var m = String(sizeKey).match(/([\d.]+)"x([\d.]+)"/);
    if (!m) { return null; }
    return { w: parseFloat(m[1]), h: parseFloat(m[2]) };
}

function buildManifestComparison(packingData, individualResults) {
    // Returns structured data for building the UI, not a text string
    var result = {
        hasData: false,
        expectedText: '',
        foundText: '',
        mismatches: [],
        notes: []
    };

    // Display names per normalized material key ("3mm ACM" and "3mmACM" are one material)
    var display = {};
    function mk(name) {
        var key = ProdName.matKey(name);
        if (!display[key]) { display[key] = name; }
        return key;
    }

    // ---- Notes that apply with or without proof data ----
    var blkFiles = {};   // key -> [labels]
    var backFiles = {};  // key -> [labels]
    if (individualResults) {
        for (var n = 0; n < individualResults.length; n++) {
            var nr = individualResults[n];
            var np = nr.parsed;
            if (np && np.type) {
                if (np.blkbk && np.type !== 'CUT') {
                    var bk = mk(np.material);
                    if (!blkFiles[bk]) { blkFiles[bk] = []; }
                    blkFiles[bk].push(ProdName.label(np) + (np.legacyBlk ? ' (old name)' : ''));
                }
                if (nr.countMode === 'back') {
                    var bb = mk(np.material);
                    if (!backFiles[bb]) { backFiles[bb] = []; }
                    backFiles[bb].push(ProdName.label(np));
                }
            }
            if (nr.runWarnings) {
                for (var w = 0; w < nr.runWarnings.length; w++) {
                    result.mismatches.push(nr.name + ': ' + nr.runWarnings[w]);
                }
            }
            if (nr.artboardNotes) {
                for (var an = 0; an < nr.artboardNotes.length; an++) {
                    result.notes.push(nr.artboardNotes[an]);
                }
            }
        }
    }
    var nk;
    for (nk in blkFiles) {
        if (blkFiles.hasOwnProperty(nk)) {
            result.notes.push('Black-back stock (counted as ' + display[nk] + '): ' + blkFiles[nk].join(', '));
        }
    }
    for (nk in backFiles) {
        if (backFiles.hasOwnProperty(nk)) {
            result.notes.push('Printed back files (' + display[nk] + '): ' + backFiles[nk].join(', '));
        }
    }

    if (!packingData) {
        return result;
    }
    result.hasData = true;

    // Build expected sizes by material from packing data
    var expectedSizes = {};
    if (packingData.materials) {
        for (var m = 0; m < packingData.materials.length; m++) {
            var material = packingData.materials[m];
            var expParsed = ProdName.parseArtboardName(material.material_name || 'Unknown');
            var materialName = mk(expParsed.material || 'Unknown');
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

    // Signs the Editor could not nest (e.g. a 48" diamond that only fits turned
    // square) are saved separately in the proof data; they are still expected.
    var unpackedCount = 0;
    if (packingData.unpacked) {
        for (var up = 0; up < packingData.unpacked.length; up++) {
            var us = packingData.unpacked[up];
            var uw = us.width || 0;
            var uh = us.height || 0;
            if (uw <= 0 || uh <= 0) { continue; }
            var uParsed = ProdName.parseArtboardName(us.material_name || 'Unknown');
            var uKey = mk(uParsed.material || 'Unknown');
            if (!expectedSizes[uKey]) { expectedSizes[uKey] = {}; }
            var uSize = normalizeDimension(uw, uh);
            expectedSizes[uKey][uSize] = (expectedSizes[uKey][uSize] || 0) + 1;
            unpackedCount++;
        }
    }
    if (unpackedCount > 0) {
        result.notes.push(unpackedCount + ' sign(s) were not nested by the Editor but are still expected');
    }

    // Build found sizes by material from file check results, with run multiples.
    // Also track which materials use CutOnly/PnC (uncountable via spot color)
    var foundSizes = {};
    var uncountableMaterials = {};
    if (individualResults) {
        for (var d = 0; d < individualResults.length; d++) {
            var ir = individualResults[d];
            var ip = ir.parsed || { type: '', material: '' };
            if (ir.countMode === 'back' || ir.countMode === 'skip') { continue; }
            var mult = (ir.runMult === undefined) ? 1 : ir.runMult;

            // Does this result carry per-artboard material buckets? (PRIME files)
            var byMat = ir.cutThroughSizesByMaterial;
            var hasByMat = false;
            if (byMat) { for (var bmk in byMat) { if (byMat.hasOwnProperty(bmk)) { hasByMat = true; break; } } }

            if (ip.type && ip.material) {
                // Production file: material from the file name
                var fileMat = mk(ip.material);
                if (!foundSizes[fileMat]) { foundSizes[fileMat] = {}; }
                if (ip.type === 'CutOnly' || ip.type === 'PnC') {
                    uncountableMaterials[fileMat] = true;
                }
                if (ir.cutThroughSizes) {
                    for (var fsize in ir.cutThroughSizes) {
                        if (ir.cutThroughSizes.hasOwnProperty(fsize)) {
                            foundSizes[fileMat][fsize] = (foundSizes[fileMat][fsize] || 0) + ir.cutThroughSizes[fsize] * mult;
                        }
                    }
                }
            } else if (hasByMat) {
                // PRIME / combined file: each cut path bucketed under the material of the
                // artboard it sits on, already multiplied by that artboard's _x#.
                for (var pm in byMat) {
                    if (!byMat.hasOwnProperty(pm)) { continue; }
                    var pk = mk(pm);
                    if (!foundSizes[pk]) { foundSizes[pk] = {}; }
                    for (var psize in byMat[pm]) {
                        if (byMat[pm].hasOwnProperty(psize)) {
                            foundSizes[pk][psize] = (foundSizes[pk][psize] || 0) + byMat[pm][psize];
                        }
                    }
                }
            } else {
                // No material info at all -> fall back to the old Unknown bucket.
                var uk = mk('Unknown');
                if (!foundSizes[uk]) { foundSizes[uk] = {}; }
                if (ir.cutThroughSizes) {
                    for (var usize in ir.cutThroughSizes) {
                        if (ir.cutThroughSizes.hasOwnProperty(usize)) {
                            foundSizes[uk][usize] = (foundSizes[uk][usize] || 0) + ir.cutThroughSizes[usize] * mult;
                        }
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

    // Diamonds: a square line (W = H, measured point to point) nested turned
    // square is cut as a smaller square. Match leftover square cuts whose
    // side x 1.414 is within 6% of the expected size (covers rounded corners).
    var diamondSizes = {}; // matKey -> { expectedSize: foundSquareSize }
    for (var dm = 0; dm < materialList.length; dm++) {
        var dKey = materialList[dm];
        if (uncountableMaterials[dKey]) { continue; }
        var dExp = expectedSizes[dKey];
        var dFnd = foundSizes[dKey];
        if (!dExp || !dFnd) { continue; }
        for (var eSize in dExp) {
            if (!dExp.hasOwnProperty(eSize)) { continue; }
            var ed = parseSizeKey(eSize);
            if (!ed || ed.w !== ed.h) { continue; }
            var deficit = dExp[eSize] - (dFnd[eSize] || 0);
            if (deficit <= 0) { continue; }
            for (var fSize in dFnd) {
                if (!dFnd.hasOwnProperty(fSize) || deficit <= 0) { continue; }
                var fd = parseSizeKey(fSize);
                if (!fd || fd.w !== fd.h || fSize === eSize) { continue; }
                var surplus = dFnd[fSize] - (dExp[fSize] || 0);
                if (surplus <= 0) { continue; }
                if (Math.abs(fd.w * Math.SQRT2 - ed.w) > ed.w * 0.06) { continue; }
                var take = Math.min(deficit, surplus);
                dFnd[fSize] -= take;
                if (dFnd[fSize] === 0) { delete dFnd[fSize]; }
                dFnd[eSize] = (dFnd[eSize] || 0) + take;
                deficit -= take;
                if (!diamondSizes[dKey]) { diamondSizes[dKey] = {}; }
                diamondSizes[dKey][eSize] = fSize;
                result.notes.push('Diamond: ' + take + 'x ' + eSize + ' ' + display[dKey] + ' matched by ' + fSize + ' square cuts');
            }
        }
    }

    var expectedLines = [];
    var foundLines = [];

    for (var m2 = 0; m2 < materialList.length; m2++) {
        var matKeyName = materialList[m2];
        var matName = display[matKeyName] || matKeyName;
        var expectedByMat = expectedSizes[matKeyName] || {};
        var foundByMat = foundSizes[matKeyName] || {};
        var diamonds = diamondSizes[matKeyName] || {};

        // Material header
        expectedLines.push(matName);
        foundLines.push(matName);

        // If this material uses CutOnly/PnC, skip counting -- show note instead
        if (uncountableMaterials[matKeyName]) {
            var expCount0 = 0;
            for (var sz0 in expectedByMat) { expCount0 += expectedByMat[sz0]; }
            if (expCount0 > 0) {
                expectedLines.push('  ' + expCount0 + ' sign(s) expected');
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
        for (var s1 = 0; s1 < sizeList.length; s1++) {
            var size1 = sizeList[s1];
            var expCount1 = expectedByMat[size1] || 0;
            var fndCount1 = foundByMat[size1] || 0;
            if (expCount1 > 0 && fndCount1 > 0 && expCount1 === fndCount1) {
                expectedLines.push('  ' + expCount1 + 'x  ' + size1);
                foundLines.push('  ' + fndCount1 + 'x  ' + size1 + (diamonds[size1] ? '  (diamond, cut ' + diamonds[size1] + ' sq)' : ''));
            }
        }

        // Mismatched items second
        for (var s2 = 0; s2 < sizeList.length; s2++) {
            var size2 = sizeList[s2];
            var expCount2 = expectedByMat[size2] || 0;
            var fndCount2 = foundByMat[size2] || 0;
            if (expCount2 === fndCount2 && expCount2 > 0) { continue; }

            if (expCount2 > 0 && fndCount2 > 0 && expCount2 !== fndCount2) {
                expectedLines.push('  ' + expCount2 + 'x  ' + size2 + '  <<<');
                foundLines.push('  ' + fndCount2 + 'x  ' + size2 + '  <<<');
                result.mismatches.push(matName + ' ' + size2 + ': expected ' + expCount2 + ', found ' + fndCount2);
            } else if (expCount2 > 0 && fndCount2 === 0) {
                expectedLines.push('  ' + expCount2 + 'x  ' + size2 + '  <<<');
                foundLines.push('  --missing--');
                result.mismatches.push('MISSING: ' + expCount2 + 'x ' + size2 + ' ' + matName);
            } else if (expCount2 === 0 && fndCount2 > 0) {
                expectedLines.push('  --not expected--');
                foundLines.push('  ' + fndCount2 + 'x  ' + size2 + '  <<<');
                result.mismatches.push('UNEXPECTED: ' + fndCount2 + 'x ' + size2 + ' ' + matName);
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
                parsed: ProdName.parseFileName(doc.name),
                materialName: detectedMaterial,
                fileType: extractFileType(doc.name),
                analysisTime: analysisTime,
                rasterCount: docResult.rasterCount,
                lowResImages: docResult.lowResImages || [],
                allImageDetails: docResult.allImageDetails || [],
                cutThroughSizes: docResult.cutThroughSizes || {},
                cutThroughSizesByMaterial: docResult.cutThroughSizesByMaterial || {},
                artboardNotes: docResult.artboardNotes || [],
                artboardRuns: docResult.artboardRuns || [],
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

        // Work out how many times each file's cuts are produced (_x# multiples,
        // shared CUT files, cuts already merged into PRINT files, printed backs)
        var nearbyNames = collectNearbyProductionNames(documents);
        computeRunMultipliers(individualDocumentResults, nearbyNames);

        // Build manifest comparison if available
        var manifestComparison = buildManifestComparison(proofManifest, individualDocumentResults);

        // Build report text
        var reportText = buildReport(documents.length, documentNames, totalTime, timingReport,
                                   totalImageCount, allLowResImages, allImageDetails, allCutThroughSizes,
                                   totalCutThroughPaths, totalCompoundPaths, allCompoundPathWarnings,
                                   includePPI, individualDocumentResults, manifestComparison, manifestStatus);

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

        // Precompute each artboard's rectangle + parsed name. Artboards follow the
        // production naming standard (e.g. "3mmACM_BlkBk_Pt7_x3"): material is the
        // base material, _x# multiplies that board's cuts, BACK boards aren't counted.
        var abInfo = [];
        var artboardNotes = [];
        var artboardRuns = [];
        try {
            for (var ab = 0; ab < doc.artboards.length; ab++) {
                var abName = doc.artboards[ab].name || '';
                var abParsed = ProdName.parseArtboardName(abName);
                var abMat = abParsed.material;
                if (abMat.length === 0) { abMat = 'Unknown'; }
                abInfo.push({
                    rect: doc.artboards[ab].artboardRect,
                    name: abName,
                    material: abMat,
                    mult: abParsed.mult,
                    isBack: abParsed.side === 'BACK',
                    blkbk: abParsed.blkbk,
                    cuts: 0
                });
            }
        } catch (eAB) {}

        // Which artboard contains a path, by its center point.
        // artboardRect and geometricBounds are both [L, T, R, B] in the same
        // point space, so scaleFactor is irrelevant to this containment test.
        function artboardForBounds(b) {
            var cx = (b[0] + b[2]) / 2;
            var cy = (b[1] + b[3]) / 2;
            for (var q = 0; q < abInfo.length; q++) {
                var r = abInfo[q].rect;
                if (cx >= r[0] && cx <= r[2] && cy <= r[1] && cy >= r[3]) {
                    return abInfo[q];
                }
            }
            return null;
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
                    artboardNotes: artboardNotes,
                    artboardRuns: artboardRuns,
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
                var abHit = artboardForBounds(allPaths[i].geometricBounds);
                totalPaths++;
                if (abHit) {
                    abHit.cuts++;
                    if (abHit.isBack) { continue; } // printed back: not counted
                }
                var mat = abHit ? abHit.material : 'Unknown';
                var abMult = abHit ? abHit.mult : 1;
                if (!cutThroughSizesByMaterial[mat]) { cutThroughSizesByMaterial[mat] = {}; }
                cutThroughSizesByMaterial[mat][size] = (cutThroughSizesByMaterial[mat][size] || 0) + abMult;
            }

            // Artboard notes for the report (multiples, black backs, printed backs)
            for (var an = 0; an < abInfo.length; an++) {
                var ai = abInfo[an];
                if (ai.mult > 1 || ai.isBack || ai.blkbk) {
                    artboardRuns.push(ai.name + ': ' + ai.cuts + ' cut(s)' +
                        (ai.mult > 1 ? ' x' + ai.mult + ' runs = ' + (ai.cuts * ai.mult) : '') +
                        (ai.blkbk ? ' [black-back stock, counted as ' + ai.material + ']' : '') +
                        (ai.isBack ? ' [printed back, not counted]' : ''));
                }
                if (ai.isBack && ai.cuts > 0) {
                    artboardNotes.push('Artboard ' + ai.name + ' is a BACK but has ' + ai.cuts + ' cut path(s)');
                }
            }

        } catch (e) {}

        return {
            cutThroughSizes: cutThroughSizes,
            cutThroughSizesByMaterial: cutThroughSizesByMaterial,
            artboardNotes: artboardNotes,
            artboardRuns: artboardRuns,
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
    var artboardNotes = cutThroughResults.artboardNotes || [];
    var artboardRuns = cutThroughResults.artboardRuns || [];
    var totalCutThroughPaths = cutThroughResults.totalCutThroughPaths;
    var compoundPathWarnings = cutThroughResults.compoundPathWarnings;
    var totalCompoundPaths = cutThroughResults.totalCompoundPaths;

    return {
        rasterCount: rasterCount,
        lowResImages: lowResImages,
        allImageDetails: allImageDetails,
        cutThroughSizes: cutThroughSizes,
        cutThroughSizesByMaterial: cutThroughSizesByMaterial,
        artboardNotes: artboardNotes,
        artboardRuns: artboardRuns,
        totalCutThroughPaths: totalCutThroughPaths,
        compoundPathWarnings: compoundPathWarnings,
        totalCompoundPaths: totalCompoundPaths
    };
}

function addNotesPanel(dialog, notes) {
    var notesPanel = dialog.add("panel", undefined, "NOTES");
    notesPanel.orientation = "column";
    notesPanel.alignChildren = "left";
    notesPanel.margins = [10, 15, 10, 8];
    var maxShow = 12;
    for (var i = 0; i < notes.length && i < maxShow; i++) {
        var line = notesPanel.add("statictext", undefined, notes[i]);
        line.graphics.foregroundColor = line.graphics.newPen(line.graphics.PenType.SOLID_COLOR, [0.2, 0.55, 0.9], 1);
    }
    if (notes.length > maxShow) {
        notesPanel.add("statictext", undefined, "(" + (notes.length - maxShow) + " more in the Full Report)");
    }
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

        // Notes (multiples, diamonds, black backs) - informational, not errors
        if (comparison.notes && comparison.notes.length > 0) {
            addNotesPanel(dialog, comparison.notes);
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
        if (comparison && comparison.notes && comparison.notes.length > 0) {
            addNotesPanel(dialog, comparison.notes);
        }
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
            var displayName = (result.parsed && result.parsed.type) ? ProdName.label(result.parsed) :
                (result.materialName && result.materialName.length > 0 ? result.materialName : result.name);
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
            if (docResult.parsed && docResult.parsed.type) {
                docReport.push("Read as: " + ProdName.label(docResult.parsed));
            }
            if (docResult.runNote) {
                docReport.push("Counting: " + docResult.runNote);
            }
            if (docResult.runWarnings && docResult.runWarnings.length > 0) {
                for (var rw = 0; rw < docResult.runWarnings.length; rw++) {
                    docReport.push("WARNING: " + docResult.runWarnings[rw]);
                }
            }
            if (docResult.artboardRuns && docResult.artboardRuns.length > 0) {
                docReport.push("Artboards:");
                for (var ar = 0; ar < docResult.artboardRuns.length; ar++) {
                    docReport.push("  " + docResult.artboardRuns[ar]);
                }
            }
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
                    includePPI, individualDocumentResults, comparison, manifestStatus) {
    var report = [];
    report.push("=== MULTI-DOCUMENT ANALYSIS SUMMARY ===");
    report.push("");

    // Proof comparison first - it is what the check is for
    if (comparison && comparison.hasData) {
        report.push("=== PROOF COMPARISON ===");
        if (comparison.mismatches.length > 0) {
            report.push("MISMATCHES:");
            for (var cm = 0; cm < comparison.mismatches.length; cm++) {
                report.push("  " + comparison.mismatches[cm]);
            }
        } else {
            report.push("All expected items found - quantities match");
        }
        report.push("");
    } else {
        report.push("Proof data: " + (manifestStatus || "No proof data found"));
        if (comparison && comparison.mismatches.length > 0) {
            for (var cw = 0; cw < comparison.mismatches.length; cw++) {
                report.push("  " + comparison.mismatches[cw]);
            }
        }
        report.push("");
    }
    if (comparison && comparison.notes && comparison.notes.length > 0) {
        report.push("NOTES:");
        for (var cn = 0; cn < comparison.notes.length; cn++) {
            report.push("  " + comparison.notes[cn]);
        }
        report.push("");
    }
    if (comparison && comparison.hasData) {
        report.push("EXPECTED (from proof):");
        report.push(comparison.expectedText);
        report.push("FOUND (in files, with multiples):");
        report.push(comparison.foundText);
    }

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
        if (docResult.parsed && docResult.parsed.type) {
            report.push("  Read as: " + ProdName.label(docResult.parsed));
        }
        if (docResult.runNote) {
            report.push("  Counting: " + docResult.runNote);
        }
        if (docResult.runWarnings) {
            for (var rw2 = 0; rw2 < docResult.runWarnings.length; rw2++) {
                report.push("  WARNING: " + docResult.runWarnings[rw2]);
            }
        }
        if (docResult.artboardRuns && docResult.artboardRuns.length > 0) {
            for (var ar2 = 0; ar2 < docResult.artboardRuns.length; ar2++) {
                report.push("  Artboard " + docResult.artboardRuns[ar2]);
            }
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
