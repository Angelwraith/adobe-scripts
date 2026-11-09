/*
@METADATA
{
  "name": "Sort Linked Files By Pt Number",
  "description": "Sort selected linked files. If the files are wrap sides (driver/front/passenger/rear/top in the name, matching Wrap Processor's naming), they're grouped by side and stacked top-to-bottom in the order Driver, Front, Passenger, Rear, Top (missing sides omitted), each row sorted left-to-right by panel number. Otherwise falls back to the original Pt# left-to-right sort. 0.25\" spacing; the leftmost/topmost of the selection is the starting position.",
  "version": "1.4",
  "target": "illustrator",
  "tags": ["sort", "linked", "files", "layout", "wrap", "processor"]
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

(function () {
    'use strict';

    // Horizontal (and vertical, in wrap mode) spacing between sorted items.
    var SPACING_INCHES = 0.25;
    var SPACING_PTS    = SPACING_INCHES * 72;

    // Matches a Pt followed by digits anywhere in the name (Pt1, Pt07, Pt123).
    var PT_REGEX = /Pt(\d+)/i;

    // Wrap side keywords, in stacking order: Driver, Front, Passenger, Rear,
    // Top. Mirrors the side detection in Wrap Processor (isWrapSideTIF) so the
    // two scripts agree on what counts as a wrap side. Case-insensitive.
    var SIDES = [
        { label: "Driver",    keys: ["driver"] },
        { label: "Front",     keys: ["front"] },
        { label: "Passenger", keys: ["passenger", "pssngr"] },
        { label: "Rear",      keys: ["rear"] },
        { label: "Top",       keys: ["top"] }
    ];

    if (app.documents.length === 0) {
        alert("Please open a document first.");
        return;
    }

    var doc = app.activeDocument;
    var sel = doc.selection;

    if (!sel || sel.length === 0) {
        alert("Select the linked files you want sorted.");
        return;
    }

    // Gather placed items with their names, preserving selection order.
    var placed = [];
    var skippedNonPlaced = 0;
    for (var i = 0; i < sel.length; i++) {
        var it = sel[i];
        if (it.typename !== "PlacedItem") { skippedNonPlaced++; continue; }
        placed.push({ item: it, name: getItemName(it), selOrder: placed.length });
    }

    // Decide mode: if 2+ placed items look like wrap sides, do wrap grouping;
    // otherwise use the original Pt# sort.
    // Production files (_PRINT_/_CUT_/_PnC_/_CutOnly_) are never wrap panels,
    // even when they carry a _FRONT/_BACK side token.
    var wrapCount = 0;
    for (var p = 0; p < placed.length; p++) {
        if (ProdName.parseFileName(placed[p].name).type !== "") continue;
        if (getSideIndex(placed[p].name) !== -1) wrapCount++;
    }

    if (wrapCount >= 2) {
        runWrapSort();
    } else {
        runPtSort();
    }

    // ------------------------------------------------------------------
    // SHARED HELPERS
    // ------------------------------------------------------------------
    function bnds(item) { return item.geometricBounds; } // [L, T, R, B]

    function getItemName(it) {
        // Prefer the linked file's name on disk; fall back to the item's
        // own name in the layers panel if there's no linked file.
        var name = "";
        try { if (it.file) name = it.file.name; } catch (e) {}
        if (!name) { try { name = it.name; } catch (e2) {} }
        try { name = decodeURI(name); } catch (e3) {}
        return name;
    }

    // Returns the index into SIDES of the first matching side keyword, or -1.
    function getSideIndex(name) {
        var lower = String(name).toLowerCase();
        for (var s = 0; s < SIDES.length; s++) {
            for (var k = 0; k < SIDES[s].keys.length; k++) {
                if (lower.indexOf(SIDES[s].keys[k]) !== -1) return s;
            }
        }
        return -1;
    }

    // Trailing panel number, e.g. "..._Pass_2" -> 2. Named panels with no
    // trailing number (e.g. "..._Rear_Bumper") return null. The extension is
    // dropped first so only the panel's own number is considered.
    function getPanelNumber(name) {
        var base = String(name).replace(/\.[^.]+$/, "");
        var m = base.match(/(\d+)\s*$/);
        return m ? parseInt(m[1], 10) : null;
    }

    // ------------------------------------------------------------------
    // WRAP MODE -- group by side, stack rows, sort each row by panel number
    // ------------------------------------------------------------------
    function runWrapSort() {
        var groups = [];
        for (var s = 0; s < SIDES.length; s++) groups[s] = [];

        var skippedNoSide = 0, skippedNames = [];

        for (var i = 0; i < placed.length; i++) {
            var si = getSideIndex(placed[i].name);
            if (si === -1) {
                skippedNoSide++;
                if (skippedNames.length < 5) skippedNames.push(placed[i].name || "(unnamed)");
                continue;
            }
            placed[i].sideIndex = si;
            placed[i].panel = getPanelNumber(placed[i].name);
            groups[si].push(placed[i]);
        }

        // Build the ordered list of rows in side order, skipping empty sides.
        // Within a side: numbered panels first (ascending by number), then any
        // named panels in selection order.
        var ordered = [];
        for (var g = 0; g < groups.length; g++) {
            var grp = groups[g];
            if (grp.length === 0) continue;

            var numbered = [], named = [];
            for (var n = 0; n < grp.length; n++) {
                if (grp[n].panel === null) named.push(grp[n]); // keep selection order
                else numbered.push(grp[n]);
            }
            numbered.sort(function (a, b) {
                if (a.panel !== b.panel) return a.panel - b.panel;
                return a.selOrder - b.selOrder; // tie-break: stable by selection
            });
            ordered.push(numbered.concat(named));
        }

        if (ordered.length === 0) {
            alert("No wrap side files found in the selection.");
            return;
        }

        // Reference: leftmost X and topmost Y across all grouped items.
        var startX = Infinity, topY = -Infinity;
        for (var r = 0; r < ordered.length; r++) {
            for (var c = 0; c < ordered[r].length; c++) {
                var b = bnds(ordered[r][c].item);
                if (b[0] < startX) startX = b[0];
                if (b[1] > topY)   topY   = b[1];
            }
        }

        // Place rows top-to-bottom, each row left-to-right at SPACING_PTS.
        var rowTop = topY;
        var totalItems = 0;
        for (var rr = 0; rr < ordered.length; rr++) {
            var cursorX = startX;
            var maxH = 0;
            for (var cc = 0; cc < ordered[rr].length; cc++) {
                var pi = ordered[rr][cc].item;
                var pb = bnds(pi);
                var w = pb[2] - pb[0];
                var h = pb[1] - pb[3];

                pi.translate(cursorX - pb[0], rowTop - pb[1]);

                cursorX += w + SPACING_PTS;
                if (h > maxH) maxH = h;
                totalItems++;
            }
            // Next row drops below the tallest item in this row, plus spacing.
            rowTop = rowTop - maxH - SPACING_PTS;
        }

        app.redraw();

        var lines = [];
        for (var li = 0; li < ordered.length; li++) {
            var sIdx = ordered[li][0].sideIndex;
            lines.push("  " + SIDES[sIdx].label + ": " + ordered[li].length + " panel(s)");
        }

        var msg = "Wrap sort complete.\n" +
                  "------------------------------\n" +
                  "Sorted " + totalItems + " panel(s) into " + ordered.length + " side row(s):\n" +
                  lines.join("\n") + "\n\n" +
                  "Stacked top-to-bottom (Driver, Front, Passenger, Rear, Top).\n" +
                  "Each row left-to-right by panel number; " + SPACING_INCHES + '" spacing.';

        if (skippedNonPlaced > 0) msg += "\n\nSkipped (not a linked/placed file): " + skippedNonPlaced;
        if (skippedNoSide > 0) {
            msg += "\n\nSkipped (no recognizable side): " + skippedNoSide;
            for (var sn = 0; sn < skippedNames.length; sn++) msg += "\n  - " + skippedNames[sn];
        }
        alert(msg);
    }

    // ------------------------------------------------------------------
    // PT# MODE -- original left-to-right sort by Pt# in the filename
    // ------------------------------------------------------------------
    function runPtSort() {
        var items = [];
        var skippedNoPt = 0, skippedNoPtNames = [];

        for (var i = 0; i < placed.length; i++) {
            // Production names: read the part with the shared parser (first part of
            // a range, _x# ignored). Anything else: first "Pt#" in the name.
            var pp = ProdName.parseFileName(placed[i].name);
            var ptNum = null;
            var group = "";
            if (pp.type !== "" && pp.parts.length > 0) {
                ptNum = pp.parts[0];
                group = ProdName.matKey(pp.material) + (pp.blkbk ? "|BLKBK" : "") + (pp.side === "BACK" ? "|BACK" : "");
            } else {
                var match = placed[i].name.match(PT_REGEX);
                if (match) { ptNum = parseInt(match[1], 10); }
            }
            if (ptNum === null) {
                skippedNoPt++;
                if (skippedNoPtNames.length < 5) skippedNoPtNames.push(placed[i].name || "(unnamed)");
                continue;
            }
            items.push({
                item: placed[i].item,
                ptNumber: ptNum,
                group: group,
                typeOrder: (pp.type === "CUT") ? 1 : 0,
                name: placed[i].name
            });
        }

        if (items.length < 2) {
            alert("Need at least 2 linked files with 'Pt#' in the filename.\n" +
                  "Found: " + items.length + " usable file(s).");
            return;
        }

        // By part number; same part: regular before BlkBk/BACK, PRINT before CUT
        items.sort(function (a, b) {
            if (a.ptNumber !== b.ptNumber) return a.ptNumber - b.ptNumber;
            if (a.group !== b.group) return a.group < b.group ? -1 : 1;
            return a.typeOrder - b.typeOrder;
        });

        var startX = Infinity, topY = -Infinity;
        for (var k = 0; k < items.length; k++) {
            var b = bnds(items[k].item);
            if (b[0] < startX) startX = b[0];
            if (b[1] > topY)   topY   = b[1];
        }

        var cursorX = startX;
        for (var j = 0; j < items.length; j++) {
            var pi = items[j].item;
            var pb = bnds(pi);
            var width = pb[2] - pb[0];
            pi.translate(cursorX - pb[0], topY - pb[1]);
            cursorX += width + SPACING_PTS;
        }

        app.redraw();

        var msg = "Sort complete.\n" +
                  "------------------------------\n" +
                  "Sorted: " + items.length + " linked file(s) by Pt#\n" +
                  "Spacing: " + SPACING_INCHES + '" between items\n' +
                  "Tops aligned to the topmost item in the selection.";

        if (skippedNonPlaced > 0) msg += "\n\nSkipped (not a linked/placed file): " + skippedNonPlaced;
        if (skippedNoPt > 0) {
            msg += "\n\nSkipped (no Pt# in filename): " + skippedNoPt;
            for (var n = 0; n < skippedNoPtNames.length; n++) msg += "\n  - " + skippedNoPtNames[n];
            if (skippedNoPt > skippedNoPtNames.length) {
                msg += "\n  ...and " + (skippedNoPt - skippedNoPtNames.length) + " more";
            }
        }
        alert(msg);
    }
})();
// end of script
