/*@METADATA{
  "name": "Open Production Files and organize for review",
  "description": "Open print and cut files, organize them together and tile the view to show them all",
  "version": "2.3",
  "target": "illustrator",
  "tags": ["review", "check", "file"]
}@END_METADATA*/

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
    // Detect OS and set Downloads path
    var downloadsPath;
    if ($.os.indexOf("Windows") !== -1) {
        // Windows path - use backslash
        var userProfile = $.getenv("USERPROFILE");
        downloadsPath = userProfile + "\\Downloads";
    } else {
        // Mac path
        downloadsPath = "~/Downloads";
    }
    
    var downloadsFolder = new Folder(downloadsPath);
    
    // Get the Filecheck Staging folder
    var stagingPath;
    if ($.os.indexOf("Windows") !== -1) {
        stagingPath = downloadsPath + "\\Filecheck Staging";
    } else {
        stagingPath = downloadsPath + "/Filecheck Staging";
    }
    
    var stagingFolder = new Folder(stagingPath);
    
    // Get list of folders in Filecheck Staging
    var subfolders = [];
    if (stagingFolder.exists) {
        var items = stagingFolder.getFiles();
        for (var i = 0; i < items.length; i++) {
            if (items[i] instanceof Folder) {
                // Ignore .stfolder folders
                if (items[i].name !== ".stfolder") {
                    subfolders.push(items[i]);
                }
            }
        }
    }

    // Load XMP library once for reading proof metadata
    try {
        if (ExternalObject.AdobeXMPScript == undefined) {
            ExternalObject.AdobeXMPScript = new ExternalObject('lib:AdobeXMPScript');
        }
        XMPMeta.registerNamespace('http://extremecolor.net/packing/', 'packing');
    } catch (e) {
        // XMP not available -- dropdown will fall back to folder names
    }

    // Find the proof file in a folder's PRIME subfolder.
    // "Prod Proof" wins over a plain "Proof" when both exist.
    function findProofFile(folder) {
        var sep = ($.os.indexOf("Windows") !== -1) ? "\\" : "/";
        var primeFolder = new Folder(folder.fsName + sep + "PRIME");
        if (!primeFolder.exists) return null;
        var primeFiles = primeFolder.getFiles();
        var proofFile = null;
        for (var j = 0; j < primeFiles.length; j++) {
            if (!(primeFiles[j] instanceof File)) continue;
            if (/^[._~]/.test(primeFiles[j].name)) continue;
            var fname = primeFiles[j].name.toUpperCase();
            if (!/\.(PDF|AI)$/.test(fname)) continue;
            if (fname.indexOf("PROOF") === -1) continue;
            if (fname.replace(/[-_\s]+/g, "").indexOf("PRODPROOF") !== -1) {
                return primeFiles[j]; // Prod Proof -- take it immediately
            }
            if (!proofFile) proofFile = primeFiles[j];
        }
        return proofFile;
    }

    // Parse an XMP/ISO 8601 date ("2026-10-14T15:22:10.123-04:00") to ms since epoch
    function parseXmpDate(str) {
        var m = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2}))?(?:\.\d+)?(Z|[+-]\d{2}:?\d{2})?)?/.exec(str);
        if (!m) return null;
        var ms = Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
        if (m[7] && m[7] !== "Z") {
            var tz = m[7].replace(":", "");
            var off = (+tz.substr(1, 2)) * 60 + (+tz.substr(3, 2));
            ms -= (tz.charAt(0) === "-" ? -1 : 1) * off * 60000;
        } else if (!m[7] && m[4]) {
            // No zone given -- treat as local time
            ms = new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +(m[6] || 0)).getTime();
        }
        return ms;
    }

    function pad2(n) { return (n < 10 ? "0" : "") + n; }
    function formatStamp(ms) {
        var d = new Date(ms);
        var h = d.getHours(), ap = h >= 12 ? "p" : "a";
        h = h % 12; if (h === 0) h = 12;
        return pad2(d.getMonth() + 1) + "/" + pad2(d.getDate()) + " " + h + ":" + pad2(d.getMinutes()) + ap;
    }

    // Build dropdown info for a folder: display name (from proof packing metadata)
    // and save timestamp (proof XMP ModifyDate -> proof file modified -> folder modified)
    function getFolderInfo(folder) {
        var info = { folder: folder, name: folder.name.replace(/%20/g, " "), time: null };
        var proofFile = null;
        try { proofFile = findProofFile(folder); } catch (e) {}

        if (proofFile) {
            try {
                var xmpFile = new XMPFile(proofFile.fsName, XMPConst.FILE_PDF, XMPConst.OPEN_FOR_READ);
                var xmp = xmpFile.getXMP();
                xmpFile.closeFile();

                if (xmp.doesPropertyExist(XMPConst.NS_XMP, 'ModifyDate')) {
                    info.time = parseXmpDate(String(xmp.getProperty(XMPConst.NS_XMP, 'ModifyDate').value));
                }

                if (xmp.doesPropertyExist('http://extremecolor.net/packing/', 'data')) {
                    var packingJSON = xmp.getProperty('http://extremecolor.net/packing/', 'data').value;
                    var packingData = eval('(' + packingJSON + ')');

                    // Build label: Client - Property - Project (omit property if not available)
                    var parts = [];
                    if (packingData.client_name) parts.push(packingData.client_name);
                    if (packingData.property_name) parts.push(packingData.property_name);
                    if (packingData.project_name) parts.push(packingData.project_name);
                    if (parts.length > 0) info.name = parts.join(" - ");
                }
            } catch (e) {
                // Any error reading metadata -- keep folder name / fall through to file dates
            }
            if (info.time === null && proofFile.modified) info.time = proofFile.modified.getTime();
        }
        if (info.time === null && folder.modified) info.time = folder.modified.getTime();
        return info;
    }

    // Read info for every folder and sort newest first
    var folderInfos = [];
    for (var i = 0; i < subfolders.length; i++) {
        folderInfos.push(getFolderInfo(subfolders[i]));
    }
    folderInfos.sort(function(a, b) {
        return (b.time || 0) - (a.time || 0);
    });
    subfolders = [];
    for (var i = 0; i < folderInfos.length; i++) {
        subfolders.push(folderInfos[i].folder);
    }

    // Create dialog window
    var dialog = new Window("dialog", "Print/Cut File Processor");
    dialog.alignChildren = "fill";
    dialog.spacing = 15;
    dialog.margins = 20;

    // Dropdown selection (if folders exist in Downloads)
    var folderDropdown = null;
    if (subfolders.length > 0) {
        var dropdownGroup = dialog.add("group");
        dropdownGroup.orientation = "column";
        dropdownGroup.alignChildren = "left";

        dropdownGroup.add("statictext", undefined, "Select folder from Downloads:");

        var dropdownRow = dropdownGroup.add("group");
        dropdownRow.orientation = "row";
        dropdownRow.spacing = 10;

        folderDropdown = dropdownRow.add("dropdownlist", undefined, []);
        folderDropdown.preferredSize.width = 400;

        for (var i = 0; i < folderInfos.length; i++) {
            var fi = folderInfos[i];
            var displayName = (fi.time ? formatStamp(fi.time) + "  |  " : "") + fi.name;
            folderDropdown.add("item", displayName);
        }
        folderDropdown.selection = 0;

        dialog.add("panel");
    }
    
    // Manual folder path input
    var pathGroup = dialog.add("group");
    pathGroup.orientation = "column";
    pathGroup.alignChildren = "left";
    
    pathGroup.add("statictext", undefined, "Or enter folder path manually:");
    var folderPathInput = pathGroup.add("edittext", undefined, "");
    folderPathInput.preferredSize.width = 400;
    
    // Buttons
    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "center";
    var openFolderButton = buttonGroup.add("button", undefined, "Open Folder");
    var okButton = buttonGroup.add("button", undefined, "Process Files", {name: "ok"});
    var cancelButton = buttonGroup.add("button", undefined, "Cancel", {name: "cancel"});
    
    // Open folder button handler
    openFolderButton.onClick = function() {
        var folderToOpen = "";
        
        // Determine which path to use
        if (folderPathInput.text != "") {
            folderToOpen = folderPathInput.text;
        } else if (folderDropdown != null && folderDropdown.selection != null) {
            var selectedIndex = folderDropdown.selection.index;
            folderToOpen = subfolders[selectedIndex].fsName;
        }
        
        if (folderToOpen == "") {
            alert("Please select a folder from the dropdown or enter a folder path first.");
            return;
        }
        
        var folder = new Folder(folderToOpen);
        if (!folder.exists) {
            alert("The specified folder does not exist.");
            return;
        }
        
        // Open folder in Windows Explorer or Mac Finder
        try {
            if ($.os.indexOf("Windows") !== -1) {
                // Windows - open in Explorer
                var batFile = new File(Folder.temp + "/openFolder.bat");
                batFile.open("w");
                batFile.writeln("@echo off");
                batFile.writeln("explorer \"" + folder.fsName + "\"");
                batFile.close();
                batFile.execute();
            } else {
                // Mac - open in Finder
                folder.execute();
            }
        } catch (e) {
            alert("Error opening folder: " + e.message);
        }
    };
    
    // Show dialog
    if (dialog.show() == 1) {
        var folderPath = "";
        
        // Determine which path to use
        if (folderPathInput.text != "") {
            // Manual path takes priority
            folderPath = folderPathInput.text;
        } else if (folderDropdown != null && folderDropdown.selection != null) {
            // Use selected dropdown folder
            var selectedIndex = folderDropdown.selection.index;
            folderPath = subfolders[selectedIndex].fsName;
        }
        
        if (folderPath == "") {
            alert("Please select a folder from the dropdown or enter a folder path.");
            return;
        }
        
        var folder = new Folder(folderPath);
        if (!folder.exists) {
            alert("The specified folder does not exist.");
            return;
        }
        
        processFiles(folder);
    }
    
    function processFiles(folder) {
        try {
            // Step 1: Find and open all PRINT, PnC, and CutOnly files
            // (file type read with the shared name parser, so "CUT" inside a
            // project name can't be mistaken for a cut file)
            var printFiles = findFiles(folder, ["PRINT", "PnC", "CutOnly"]);

            // Open all print files
            var openDocs = [];
            for (var i = 0; i < printFiles.length; i++) {
                try {
                    var doc = app.open(printFiles[i]);
                    var printParsed = ProdName.parseFileName(safeFileName(printFiles[i]));
                    openDocs.push({
                        doc: doc,
                        file: printFiles[i],
                        parsed: printParsed,
                        isCutOnly: printParsed.type === "CutOnly"
                    });
                } catch (e) {
                    alert("Error opening file: " + printFiles[i].name + "\n" + e.message);
                }
            }

            // Step 2: Find all CUT files
            var cutFiles = findFiles(folder, ["CUT"]);

            // If nothing was found at all, alert and bail
            if (printFiles.length == 0 && cutFiles.length == 0) {
                alert("No PRINT, PnC, CutOnly, or CUT files found.");
                return;
            }

            // Check if any opened docs are non-CutOnly (i.e. PRINT/PnC that need CUT placed)
            var hasNonCutOnly = false;
            for (var i = 0; i < openDocs.length; i++) {
                if (!openDocs[i].isCutOnly) {
                    hasNonCutOnly = true;
                    break;
                }
            }

            // Step 3: Match CUT files to PRINT docs, place matched ones, open unmatched ones standalone.
            // Pairing ignores _x# (PRINT_..._Pt8_x2 pairs with CUT_..._Pt8 or CUT_..._Pt8_x2)
            // and a shared CUT (CUT_..._Pt3-5and8) is placed into every PRINT it serves.
            // BACK print files never receive cuts.
            var printCount = openDocs.length;
            // Every exact cut key, so a _FRONT print only falls back to an untagged
            // CUT when no _FRONT cut exists for it.
            var allCutKeys = {};
            for (var ak = 0; ak < cutFiles.length; ak++) {
                var akP = ProdName.parseFileName(safeFileName(cutFiles[ak]));
                var akParts = akP.parts.length ? akP.parts : [0];
                for (var akq = 0; akq < akParts.length; akq++) { allCutKeys[ProdName.pairKey(akP, akParts[akq])] = true; }
            }
            for (var i = 0; i < cutFiles.length; i++) {
                var cutFile = cutFiles[i];
                var cutParsed = ProdName.parseFileName(safeFileName(cutFile));
                var cutParts = cutParsed.parts.length ? cutParsed.parts : [0];
                var cutKeys = {};
                for (var ck = 0; ck < cutParts.length; ck++) {
                    cutKeys[ProdName.pairKey(cutParsed, cutParts[ck])] = true;
                }

                // Find matching PRINT/PnC files (skip CutOnly files - they don't need a separate CUT file placed)
                var matchedDocs = [];
                if (hasNonCutOnly) {
                    for (var j = 0; j < printCount; j++) {
                        if (openDocs[j].isCutOnly) { continue; }
                        var dp = openDocs[j].parsed;
                        if (cutKeys[ProdName.pairKey(dp)]) {
                            matchedDocs.push(openDocs[j].doc);
                        } else if (dp.side === 'FRONT' && !cutParsed.side &&
                                   !allCutKeys[ProdName.pairKey(dp)] && cutKeys[ProdName.pairKeyNoSide(dp)]) {
                            matchedDocs.push(openDocs[j].doc);
                        }
                    }
                }

                if (matchedDocs.length > 0) {
                    // Place CUT into each matching PRINT doc
                    for (var md = 0; md < matchedDocs.length; md++) {
                        placeCutFile(matchedDocs[md], cutFile);
                    }
                } else {
                    // No matching PRINT file - open the CUT file as a standalone document
                    try {
                        var cutDoc = app.open(cutFile);
                        openDocs.push({
                            doc: cutDoc,
                            file: cutFile,
                            parsed: cutParsed,
                            isCutOnly: false
                        });
                    } catch (e) {
                        alert("Error opening CUT file: " + cutFile.name + "\n" + e.message);
                    }
                }
            }

            // Arrange documents in a tiled grid
            app.executeMenuCommand("tile");

            // Fit all documents in their windows
            for (var i = 0; i < openDocs.length; i++) {
                app.activeDocument = openDocs[i].doc;
                app.executeMenuCommand("fitall");
            }
                  
        } catch (e) {
            alert("Error: " + e.message);
        }
    }
    
    function findFiles(folder, keywords) {
        var foundFiles = [];
        var files = folder.getFiles();
        
        for (var i = 0; i < files.length; i++) {
            if (files[i] instanceof Folder) {
                // Recursively search subfolders
                var subFiles = findFiles(files[i], keywords);
                foundFiles = foundFiles.concat(subFiles);
            } else if (files[i] instanceof File) {
                var fileName = files[i].name.toUpperCase();

                // Skip hidden/temp/partial files (e.g. "._Foo.ai", "~syncthing~Foo.ai.tmp")
                if (/^[._~]/.test(files[i].name)) {
                    continue;
                }

                // File must actually END with .ai/.eps/.pdf (not just contain it,
                // which would match things like "Foo_PRINT.ai.tmp" or "Foo_CUT.pdf.crdownload"
                // and cause Illustrator to open them as text)
                if (!/\.(AI|EPS|PDF)$/.test(fileName)) {
                    continue;
                }

                // Match on the file TYPE token (_PRINT_, _CUT_, _PnC_, _CutOnly_)
                var fileType = ProdName.parseFileName(safeFileName(files[i])).type.toUpperCase();
                for (var k = 0; k < keywords.length; k++) {
                    if (fileType !== "" && fileType === keywords[k].toUpperCase()) {
                        foundFiles.push(files[i]);
                        break;
                    }
                }
            }
        }
        
        return foundFiles;
    }
    
    function safeFileName(f) {
        try { return decodeURI(f.name); } catch (e) { return String(f.name); }
    }

    function getBaseName(fileName) {
        // Remove keywords and file extension to get base name for matching
        var baseName = fileName.toUpperCase();
        baseName = baseName.replace(/_PRINT_/gi, "_");
        baseName = baseName.replace(/_PNC_/gi, "_");
        baseName = baseName.replace(/_CUTONLY_/gi, "_");
        baseName = baseName.replace(/_CUT_/gi, "_");
        baseName = baseName.replace(/\.AI$/gi, "");
        baseName = baseName.replace(/\.EPS$/gi, "");
        baseName = baseName.replace(/\.PDF$/gi, "");
        baseName = baseName.replace(/[-_\s]+/g, ""); // Remove separators
        baseName = baseName.replace(/^\s+|\s+$/g, ""); // Manual trim for ExtendScript
        
        return baseName;
    }
    
    function placeCutFile(doc, cutFile) {
        try {
            app.activeDocument = doc;
            
            // Get artboard dimensions
            var artboard = doc.artboards[0];
            var artboardRect = artboard.artboardRect;
            var artboardWidth = artboardRect[2] - artboardRect[0];
            var artboardHeight = artboardRect[1] - artboardRect[3];
            
            // Place the CUT file
            var placedItem = doc.placedItems.add();
            placedItem.file = cutFile;
            
            // Check if placed item is 1/10th the size of artboard (the Illustrator bug)
            var placedWidth = placedItem.width;
            var placedHeight = placedItem.height;
            
            var widthRatio = artboardWidth / placedWidth;
            var heightRatio = artboardHeight / placedHeight;
            
            // If either dimension is approximately 10x larger, scale up
            if ((widthRatio > 8 && widthRatio < 12) || (heightRatio > 8 && heightRatio < 12)) {
                placedItem.resize(1000, 1000); // Scale by 10x (1000%)
            }
            
            // Center the placed item on the artboard
            var itemWidth = placedItem.width;
            var itemHeight = placedItem.height;
            
            var centerX = artboardRect[0] + (artboardWidth / 2);
            var centerY = artboardRect[1] - (artboardHeight / 2);
            
            placedItem.left = centerX - (itemWidth / 2);
            placedItem.top = centerY + (itemHeight / 2);
            
            // Embed the placed item
            placedItem.embed();
            
        } catch (e) {
            alert("Error placing CUT file " + cutFile.name + ": " + e.message);
        }
    }
})();