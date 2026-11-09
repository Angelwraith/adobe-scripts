/*
@METADATA
{
  "name": "Link Namer",
  "description": "Name Linked Files for ProdProof Layout",
  "version": "1.5",
  "target": "illustrator",
  "tags": ["ProdProof", "linked", "namer"]
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
    if (!app.documents.length) {
        alert("Please open a document first.");
        return;
    }
    
    var doc = app.activeDocument;
    var selection = doc.selection;
    
    if (selection.length === 0) {
        alert("Please select one or more linked files.");
        return;
    }
    
    // Get the host file name (without extension)
    var hostFileName = doc.name;
    var dotIndex = hostFileName.lastIndexOf(".");
    if (dotIndex > -1) {
        hostFileName = hostFileName.substring(0, dotIndex);
    }
    
    var processedCount = 0;
    
    for (var i = 0; i < selection.length; i++) {
        var item = selection[i];
        if (item.typename === "PlacedItem") {
            processLinkedFile(item, hostFileName);
            processedCount++;
        }
    }
    
    if (processedCount === 0) {
        alert("No linked files found in selection.");
        return;
    }
    
    // Run CUT Mover silently after processing
    runCutMover();
    
    function processLinkedFile(placedItem, hostFileName) {
        try {
            var linkedFileName = placedItem.file.name;
            
            // Remove file extension
            var dotIndex = linkedFileName.lastIndexOf(".");
            if (dotIndex > -1) {
                linkedFileName = linkedFileName.substring(0, dotIndex);
            }
            
            // Replace %20 with spaces (URL decoding)
            linkedFileName = linkedFileName.replace(/%20/g, " ");
            
            // Remove the host file prefix to isolate the material/part info
            var hostPrefix = hostFileName;
            
            // Remove "Proof", "PRIME", "ProdProof" etc. from host name to get base prefix
            hostPrefix = hostPrefix.replace(/_?(Proof|PRIME|ProdProof)$/i, "");
            
            var processedName = linkedFileName;
            
            // If the linked file starts with the host prefix, remove it
            if (linkedFileName.indexOf(hostPrefix) === 0) {
                processedName = linkedFileName.substring(hostPrefix.length);
                // Remove leading underscore if present
                if (processedName.charAt(0) === "_") {
                    processedName = processedName.substring(1);
                }
            }
            
            // Remove common keywords like CUT_, PRINT_ if they're at the beginning
            processedName = processedName.replace(/^(CUT_|PRINT_)/i, "");
            
            // Get the bounds - store them immediately
            var bounds = placedItem.geometricBounds;
            var left = bounds[0];
            var top = bounds[1]; 
            var right = bounds[2];
            var bottom = bounds[3];
            
            var placedItemWidth = Math.abs(right - left);
            var placedItemHeight = Math.abs(top - bottom);
            var widthInches = (placedItemWidth / 72) * 10;  // Multiply by 10
            var heightInches = (placedItemHeight / 72) * 10;  // Multiply by 10
            var scaledWidth = Math.round(widthInches);  // Round to whole numbers
            var scaledHeight = Math.round(heightInches);  // Round to whole numbers
            var fullName = scaledWidth + '"x' + scaledHeight + '" - ' + processedName;
            
            // Create rectangle using a completely different approach
            // Create a basic rectangle first, then directly set its properties
            
            var targetLeft = placedItem.left;
            var targetTop = placedItem.top;
            var targetWidth = placedItem.width;
            var targetHeight = placedItem.height;
            
            // Create a basic rectangle (any size)
            var whiteRect = doc.pathItems.rectangle(0, 0, -100, 100);
            
            // Now directly set its position and size properties
            whiteRect.left = targetLeft;
            whiteRect.top = targetTop;
            whiteRect.width = targetWidth;
            whiteRect.height = targetHeight;
            
            whiteRect.fillColor = createWhiteColor();
            whiteRect.filled = true;
            whiteRect.stroked = false;
            

            
            // Move rectangle behind placed item
            whiteRect.move(placedItem, ElementPlacement.PLACEAFTER);
            
            // Create text
            var textFrame = doc.textFrames.add();
            textFrame.contents = fullName;
            
            var textRange = textFrame.textRange;
            try {
                textRange.characterAttributes.textFont = app.textFonts.getByName("MyriadPro-Black");
            } catch (e) {
                try {
                    textRange.characterAttributes.textFont = app.textFonts.getByName("MyriadPro-Bold");
                } catch (e2) {
                    textRange.characterAttributes.textFont = app.textFonts[0]; 
                }
            }
            textRange.characterAttributes.size = 13;
            textRange.characterAttributes.fillColor = createBlackColor();
            textRange.paragraphAttributes.justification = Justification.CENTER;
            
            // Position text above
            var centerX = left + (placedItemWidth / 2);
            textFrame.left = centerX - (textFrame.width / 2);
            textFrame.top = top + 25;
            
            // Simple grouping - create group first
            var group = doc.groupItems.add();
            whiteRect.moveToBeginning(group);
            placedItem.moveToBeginning(group);
            textFrame.moveToBeginning(group);
            
        } catch (error) {
            alert("Error processing " + linkedFileName + ": " + error.toString());
        }
    }
    
    function extractProcessedName(linkedFileName, hostFileName) {
        var result = linkedFileName;
        var linkedDot = result.lastIndexOf(".");
        if (linkedDot > -1) {
            result = result.substring(0, linkedDot);
        }
        
        var hostNameBase = hostFileName;
        var prefixToRemove = "";

        var proofRegex = /(Proof|ProdProof)$/i;
        var match = hostNameBase.match(proofRegex);
        
        if (match) {
            var endIndex = match.index;
            prefixToRemove = hostNameBase.substring(0, endIndex);
        }

        var rawPrefixParts = prefixToRemove.split(/[_-\s]+/);
        var prefixParts = [];
        for (var k = 0; k < rawPrefixParts.length; k++) {
            if (rawPrefixParts[k].length > 0) {
                prefixParts.push(rawPrefixParts[k]);
            }
        }
        
        for (var i = 0; i < prefixParts.length; i++) {
            if (prefixParts[i].length > 0) {
                var partRegex = new RegExp("(^|[_-\\s])" + escapeRegExp(prefixParts[i]) + "([_-\\s]|$)", 'gi');
                result = result.replace(partRegex, "$1");
            }
        }

        var keywords = ["CUT", "PRINT", "PNC", "PRIME"];
        for (var j = 0; j < keywords.length; j++) {
            var keywordRegex = new RegExp("(^|[_-\\s])" + escapeRegExp(keywords[j]) + "([_-\\s]|$)", 'gi');
            result = result.replace(keywordRegex, "$1");
        }
        
        result = result.replace(/[_-\s]+/g, "_").replace(/^_|_$/g, "");
        
        return result || "Unknown";
    }

    function escapeRegExp(string) {
        return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }
    
    function createWhiteColor() {
        var color = new RGBColor();
        color.red = 255;
        color.green = 255;
        color.blue = 255;
        return color;
    }
    
    function createBlackColor() {
        var color = new RGBColor();
        color.red = 0;
        color.green = 0;
        color.blue = 0;
        return color;
    }
    
    function runCutMover() {
        var doc = app.activeDocument;
        var placedItems = [];
        var printItems = [];
        var cutItems = [];
        var matchedPairs = [];

        function getFileNameWithoutExtension(filePath) {
            var fileName = filePath.substring(filePath.lastIndexOf('/') + 1);
            fileName = fileName.substring(fileName.lastIndexOf('\\') + 1);
            return fileName.substring(0, fileName.lastIndexOf('.'));
        }

        function getFileNameParts(fileName) {
            var printMatch = fileName.match(/^(.*)_PRINT_(.*)$/i);
            if (printMatch) {
                return {
                    before: printMatch[1],
                    after: printMatch[2],
                    type: 'PRINT'
                };
            }
            
            var cutMatch = fileName.match(/^(.*)_CUT_(.*)$/i);
            if (cutMatch) {
                return {
                    before: cutMatch[1],
                    after: cutMatch[2],
                    type: 'CUT'
                };
            }
            
            return null;
        }

        function getNormalizedName(fileName) {
            var parts = getFileNameParts(fileName);
            if (parts) {
                return parts.before + "_" + parts.after;
            }
            return fileName;
        }

        function containsPrint(fileName) {
            var t = ProdName.parseFileName(fileName).type;
            return t === 'PRINT' || t === 'PnC';
        }

        function containsCut(fileName) {
            return ProdName.parseFileName(fileName).type === 'CUT';
        }

        function collectPlacedItems(container) {
            for (var i = 0; i < container.placedItems.length; i++) {
                placedItems.push(container.placedItems[i]);
            }
            
            for (var j = 0; j < container.groupItems.length; j++) {
                collectPlacedItems(container.groupItems[j]);
            }
        }

        collectPlacedItems(doc);

        if (placedItems.length === 0) {
            return;
        }

        for (var i = 0; i < placedItems.length; i++) {
            var item = placedItems[i];
            try {
                var fileName = getFileNameWithoutExtension(item.file.fsName);
                
                if (containsPrint(fileName)) {
                    printItems.push({
                        item: item,
                        fileName: fileName,
                        normalizedName: getNormalizedName(fileName)
                    });
                } else if (containsCut(fileName)) {
                    cutItems.push({
                        item: item,
                        fileName: fileName,
                        normalizedName: getNormalizedName(fileName)
                    });
                }
            } catch (e) {
                continue;
            }
        }

        // Pair with the shared name parser: _x# is ignored, BACK prints get no
        // cut, and a shared CUT (Pt3-5and8) is matched to every PRINT it serves
        // (duplicated for the 2nd, 3rd... print).
        var cutUsed = {};
        function findCut(printKey, noSideOnly) {
            for (var j = 0; j < cutItems.length; j++) {
                var cp = ProdName.parseFileName(cutItems[j].fileName);
                if (noSideOnly && cp.side) { continue; }
                var cParts = cp.parts.length ? cp.parts : [0];
                for (var q = 0; q < cParts.length; q++) {
                    if (ProdName.pairKey(cp, cParts[q]) === printKey) { return j; }
                }
            }
            return -1;
        }
        for (var i = 0; i < printItems.length; i++) {
            var printItem = printItems[i];
            var pp = ProdName.parseFileName(printItem.fileName);
            // Exact (FRONT with FRONT, one-sided with one-sided); an older _FRONT
            // print with an untagged CUT falls back to it.
            var j = findCut(ProdName.pairKey(pp), false);
            if (j < 0 && pp.side === 'FRONT') { j = findCut(ProdName.pairKeyNoSide(pp), true); }
            if (j >= 0) {
                matchedPairs.push({
                    print: printItem,
                    cut: cutItems[j],
                    duplicate: !!cutUsed[j]
                });
                cutUsed[j] = true;
            }
        }

        if (matchedPairs.length === 0) {
            return;
        }

        for (var i = 0; i < matchedPairs.length; i++) {
            var pair = matchedPairs[i];
            var printItem = pair.print.item;
            var cutItem = pair.cut.item;
            if (pair.duplicate) {
                try { cutItem = cutItem.duplicate(); } catch (eDup) { continue; }
            }

            try {
                var targetPosition = [printItem.position[0], printItem.position[1]];

                // Move the cut into the same parent as the print (its group,
                // created earlier in processLinkedFile), placed just above
                // the print in z-order. This pulls the cut into the print's
                // group automatically so the user doesn't have to group them
                // by hand afterward. PLACEBEFORE = earlier in the parent's
                // children = higher in the visual stack.
                cutItem.move(printItem, ElementPlacement.PLACEBEFORE);
                cutItem.position = targetPosition;
            } catch (e) {
                continue;
            }
        }
    }
})();
