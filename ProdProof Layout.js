/*
@METADATA
{
  "name": "ProdProof Layout",
  "description": "Finds the production files the PRIME exporter made, places them on the ProdProof production pages at a reduced scale, grouped by material in generated frames, with size labels, CUT overlays and multiples badges",
  "version": "1.3",
  "target": "illustrator",
  "tags": ["ProdProof", "production", "layout", "frames", "linked"]
}
@END_METADATA
*/

#target illustrator

// How to use: open the ProdProof and run the script. New artboards are added
// straight after the last cover page (the page with "Please see proof(s) on
// following page(s)"); existing pages are never drawn on. It links the
// production files (nothing
// in the order folder is changed), sizes them all at one scale, groups them
// by material and builds the frames around the art. Everything it creates is
// named "PPL ..." in the Layers panel.
//
// By hand afterwards: VHB / Black Back tags, color match swatches,
// reflective / laminate flags, external production table, job notes.

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

// ==== PPL MODEL ====
// Turns production file names + sizes into material groups (no Illustrator
// objects). Uses the shared ProdName parser above.
var PPLModel = (function () {
    var FILE_SCALE = 10;   // production files are 1/10 scale

    // Older exports write the run count as "_3ea" in front of the part
    // ("3mmACM_3ea_Pt16", "Alumigraphics Vinyl_2ea"); current ones use "_x3".
    // The shared parser leaves "_3ea" in the material, so lift it out here.
    function normalize(p) {
        var m = String(p.material).match(/^(.*\S)[_ ]+(\d+)\s*ea$/i);
        if (m) {
            p.material = m[1];
            if (!p.hasMult) { p.mult = Math.max(1, parseInt(m[2], 10)); p.hasMult = true; }
        }
        return p;
    }

    function displayMaterial(material) {
        return String(material).replace(/(\d)mm(?=[A-Za-z])/g, '$1MM ').toUpperCase();
    }

    function isRigid(material) {
        return /ACM|PVC|COROPLAST|ACRYLIC|FOAM|DIBOND|STYRENE|MAGNET|\d\s*MM|\b[01]?\.?\d{2,3}\s*ALUM|ALUMINUM$/i.test(material) &&
               !/VINYL|BANNER|ALUMIGRAPHICS|SATIN ALUMINUM/i.test(material);
    }

    function inches(pts) { return Math.round(pts / 72 * FILE_SCALE); }

    function rangesText(list) { return ProdName.formatParts(list); }

    // files: [{ name, path, w, h }] with w/h in points at 100%
    function build(files) {
        var prints = [];
        var cuts = [];
        var skipped = [];
        var i, j;
        for (i = 0; i < files.length; i++) {
            var p = normalize(ProdName.parseFileName(files[i].name));
            p.path = files[i].path; p.w = files[i].w; p.h = files[i].h; p.fileName = files[i].name;
            if (p.type === 'CUT') { cuts.push(p); }
            else if (p.type === 'PRINT' || p.type === 'PnC' || p.type === 'CutOnly') { prints.push(p); }
            else { skipped.push(files[i].name); }
        }

        var groups = [];
        var byKey = {};
        var warnings = [];
        for (i = 0; i < prints.length; i++) {
            var s = prints[i];
            var cls = (s.type === 'PnC') ? 'PnC' : (s.type === 'CutOnly' ? 'CutOnly' : 'PRINT');
            var key = ProdName.matKey(s.material) + '|' + (s.blkbk ? 'B' : '') + '|' + cls;
            var g = byKey[key];
            if (!g) {
                var header = displayMaterial(s.material);
                if (s.blkbk) { header = 'BLACK BACK ' + header; }
                if (cls === 'PnC') { header += ' PNC'; }
                if (cls === 'CutOnly') { header += ' CUT'; }
                g = {
                    key: key, material: s.material, header: header, sheets: [], cols: 0, area: 0,
                    style: s.blkbk ? 'blkbk' : (cls === 'PnC' ? 'pnc' : (isRigid(s.material) ? 'rigid' : 'other'))
                };
                byKey[key] = g;
                groups.push(g);
            }
            var desc = '';
            if (s.parts.length) { desc = 'Pt' + s.partsText; }
            if (s.side) { desc = s.side + (desc ? '_' + desc : ''); g.cols = 2; }
            s.label = inches(s.w) + '"x' + inches(s.h) + '"' + (desc ? ' - ' + desc : '');
            s.cut = null;
            g.sheets.push(s);
            g.area += s.w * s.h;
        }

        // Pair CUT files (a shared cut such as Pt1-3 serves every print it lists)
        function findCut(printKey, noSideOnly) {
            for (var c = 0; c < cuts.length; c++) {
                if (noSideOnly && cuts[c].side) { continue; }
                var cp = cuts[c].parts.length ? cuts[c].parts : [0];
                for (var q = 0; q < cp.length; q++) {
                    if (ProdName.pairKey(cuts[c], cp[q]) === printKey) { return c; }
                }
            }
            return -1;
        }
        var cutUsed = {};
        for (i = 0; i < prints.length; i++) {
            var pr = prints[i];
            if (pr.type !== 'PRINT' || pr.side === 'BACK') { continue; }
            var idx = findCut(ProdName.pairKey(pr), false);
            if (idx < 0 && pr.side === 'FRONT') { idx = findCut(ProdName.pairKeyNoSide(pr), true); }
            if (idx >= 0) { pr.cut = cuts[idx]; cutUsed[idx] = true; }
        }
        for (i = 0; i < cuts.length; i++) {
            if (!cutUsed[i]) { warnings.push('CUT file with no matching PRINT: ' + cuts[i].fileName); }
        }

        // Sort sheets, flag multiples and missing parts
        var multiples = [];
        for (i = 0; i < groups.length; i++) {
            var grp = groups[i];
            grp.sheets.sort(function (a, b) {
                var pa = a.parts.length ? a.parts[0] : 0;
                var pb = b.parts.length ? b.parts[0] : 0;
                if (pa !== pb) { return pa - pb; }
                var sa = a.side === 'BACK' ? 1 : 0;
                var sb = b.side === 'BACK' ? 1 : 0;
                if (sa !== sb) { return sa - sb; }
                return a.fileName < b.fileName ? -1 : (a.fileName > b.fileName ? 1 : 0);
            });
            var seen = {};
            var maxPt = 0;
            for (j = 0; j < grp.sheets.length; j++) {
                var sh = grp.sheets[j];
                if (sh.mult > 1) { multiples.push(sh.label + ' (' + grp.header + ') x' + sh.mult); }
                for (var k = 0; k < sh.parts.length; k++) {
                    seen[sh.parts[k]] = true;
                    if (sh.parts[k] > maxPt) { maxPt = sh.parts[k]; }
                }
            }
            var missing = [];
            for (j = 1; j < maxPt; j++) { if (!seen[j]) { missing.push(j); } }
            if (missing.length) {
                grp.note = 'CHECK: no Pt ' + rangesText(missing) + ' file found';
                warnings.push(grp.header + ': no Pt ' + rangesText(missing) + ' file found');
            }
        }

        // Rigid stock first, then everything else, biggest first
        groups.sort(function (a, b) {
            var ra = (a.style === 'rigid') ? 0 : 1;
            var rb = (b.style === 'rigid') ? 0 : 1;
            if (ra !== rb) { return ra - rb; }
            return b.area - a.area;
        });

        return { groups: groups, sheetCount: prints.length, cutCount: cuts.length,
                 multiples: multiples, warnings: warnings, skipped: skipped };
    }

    return { build: build, normalize: normalize, FILE_SCALE: FILE_SCALE };
})();
// ==== END PPL MODEL ====

// ==== PPL LAYOUT CORE ====
// Pure layout maths (no Illustrator objects) so it can be tested outside
// Illustrator. All sizes are points. A "sheet" is one PRINT / PnC / CutOnly
// file: { w, h } is the placed file at 100% (the files are 1/10 scale).
var PPL = (function () {
    var CFG = {
        pageMargin: 18,        // clear edge around the page
        frameGap: 8,           // between material frames
        headerH: 23,           // coloured material header strip
        padX: 20,              // frame side padding
        padTop: 8,             // header to first label / badge
        padBottom: 14,         // last sheet to frame bottom
        labelH: 15,            // 13pt label line
        labelGap: 10,          // label to sheet
        badgeH: 24,            // "x4" badge
        badgeGap: 5,           // badge to label
        cellGapX: 22,          // between sheets in a row
        rowGap: 16,            // between rows
        charW: 6.4,            // average 13pt MyriadPro-Black character
        headerCharW: 6.6,      // average 13pt header character
        noteH: 14,             // missing-part note line
        baseFraction: 0.30,    // house standard: 30% of the 1/10 file (1:33)
        fractions: [0.40, 0.35, 0.30, 0.25, 0.20, 0.15, 0.12, 0.10, 0.08, 0.06, 0.05, 0.04, 0.03]
    };

    function cellWidth(sheet, f) {
        var lw = sheet.label.length * CFG.charW;
        var sw = sheet.w * f;
        return Math.max(sw, lw);
    }

    // Break one material group into rows that fit availW, then into chunks
    // (frames) that fit availH. Returns an array of frames.
    function buildFrames(group, f, availW, availH) {
        var innerW = availW - 2 * CFG.padX;
        var headerW = group.header.length * CFG.headerCharW + 16;
        var rows = [];
        var cur = null;
        var i;
        for (i = 0; i < group.sheets.length; i++) {
            var s = group.sheets[i];
            var cw = cellWidth(s, f);
            var forceBreak = group.cols && cur && cur.cells.length >= group.cols;
            if (!cur || forceBreak || cur.w + CFG.cellGapX + cw > innerW) {
                cur = { cells: [], w: 0, sheetH: 0, hasBadge: false };
                rows.push(cur);
            }
            cur.cells.push({ sheet: s, cw: cw });
            cur.w += (cur.cells.length > 1 ? CFG.cellGapX : 0) + cw;
            if (s.h * f > cur.sheetH) { cur.sheetH = s.h * f; }
            if (s.mult > 1) { cur.hasBadge = true; }
        }
        for (i = 0; i < rows.length; i++) {
            rows[i].h = (rows[i].hasBadge ? CFG.badgeH + CFG.badgeGap : 0) +
                        CFG.labelH + CFG.labelGap + rows[i].sheetH;
        }
        var fixed = CFG.headerH + CFG.padTop + CFG.padBottom;
        var frames = [];
        var fr = null;
        for (i = 0; i < rows.length; i++) {
            var r = rows[i];
            if (!fr || fr.h + CFG.rowGap + r.h > availH) {
                fr = { group: group, rows: [], w: 0, h: fixed, part: frames.length };
                frames.push(fr);
            }
            fr.h += (fr.rows.length ? CFG.rowGap : 0) + r.h;
            fr.rows.push(r);
            if (r.w > fr.w) { fr.w = r.w; }
        }
        for (i = 0; i < frames.length; i++) {
            frames[i].w = Math.max(frames[i].w + 2 * CFG.padX, headerW);
            frames[i].fits = frames[i].w <= availW + 0.01 && frames[i].h <= availH + 0.01;
        }
        // The missing-part note sits under the last frame of the group
        if (group.note && frames.length) {
            frames[frames.length - 1].h += CFG.noteH;
            frames[frames.length - 1].note = group.note;
            var nw = group.note.length * 4.6 + 2 * CFG.padX;
            if (nw > frames[frames.length - 1].w) { frames[frames.length - 1].w = Math.min(nw, availW); }
        }
        return frames;
    }

    // Shelf packing with stacked columns: frames go left to right along a
    // shelf; a short frame may stack under another in the same column.
    function packFrames(frames, availW, availH) {
        var pages = [];
        function tryPage(pg, fr) {
            var s, c;
            for (s = 0; s < pg.shelves.length; s++) {
                var sh = pg.shelves[s];
                for (c = 0; c < sh.cols.length; c++) {
                    var col = sh.cols[c];
                    if (fr.w <= col.w + 0.01 && col.used + CFG.frameGap + fr.h <= sh.h + 0.01) {
                        fr.x = col.x; fr.y = sh.y + col.used + CFG.frameGap;
                        col.used += CFG.frameGap + fr.h;
                        return true;
                    }
                }
                var isLast = (s === pg.shelves.length - 1);
                var limitH = isLast ? (availH - sh.y) : sh.h;
                if (sh.xUsed + (sh.cols.length ? CFG.frameGap : 0) + fr.w <= availW + 0.01 && fr.h <= limitH + 0.01) {
                    var x = sh.xUsed + (sh.cols.length ? CFG.frameGap : 0);
                    fr.x = x; fr.y = sh.y;
                    sh.cols.push({ x: x, w: fr.w, used: fr.h });
                    sh.xUsed = x + fr.w;
                    if (fr.h > sh.h) { sh.h = fr.h; }
                    return true;
                }
            }
            var y = 0;
            if (pg.shelves.length) {
                var last = pg.shelves[pg.shelves.length - 1];
                y = last.y + last.h + CFG.frameGap;
            }
            if (y + fr.h <= availH + 0.01 && fr.w <= availW + 0.01) {
                pg.shelves.push({ y: y, h: fr.h, xUsed: fr.w, cols: [{ x: 0, w: fr.w, used: fr.h }] });
                fr.x = 0; fr.y = y;
                return true;
            }
            return false;
        }
        for (var i = 0; i < frames.length; i++) {
            var placed = false;
            for (var p = 0; p < pages.length && !placed; p++) {
                if (tryPage(pages[p], frames[i])) { frames[i].page = p; placed = true; }
            }
            if (!placed) {
                var pg = { shelves: [] };
                pages.push(pg);
                tryPage(pg, frames[i]);          // oversize frames still land at 0,0
                if (frames[i].x === undefined) { frames[i].x = 0; frames[i].y = 0; }
                frames[i].page = pages.length - 1;
            }
        }
        return pages.length;
    }

    function layoutAt(groups, f, pageW, pageH) {
        var availW = pageW - 2 * CFG.pageMargin;
        var availH = pageH - 2 * CFG.pageMargin;
        var frames = [];
        var fits = true;
        var g, i;
        for (g = 0; g < groups.length; g++) {
            var fs = buildFrames(groups[g], f, availW, availH);
            for (i = 0; i < fs.length; i++) {
                if (!fs[i].fits) { fits = false; }
                frames.push(fs[i]);
            }
        }
        var pageCount = packFrames(frames, availW, availH);
        // Centre each page's content horizontally, keep it at the top margin
        for (var p = 0; p < pageCount; p++) {
            var maxR = 0;
            for (i = 0; i < frames.length; i++) {
                if (frames[i].page === p && frames[i].x + frames[i].w > maxR) { maxR = frames[i].x + frames[i].w; }
            }
            var dx = CFG.pageMargin + Math.max(0, (availW - maxR) / 2);
            for (i = 0; i < frames.length; i++) {
                if (frames[i].page === p) { frames[i].x += dx; frames[i].y += CFG.pageMargin; }
            }
        }
        return { fraction: f, frames: frames, pageCount: pageCount, fits: fits };
    }

    // Auto scale: start from the house 30%. If something cannot fit a page at
    // 30%, step down until it does. Otherwise use the largest scale that does
    // not need more pages than 30% would.
    function chooseLayout(groups, pageW, pageH, forcedFraction) {
        if (forcedFraction) { return layoutAt(groups, forcedFraction, pageW, pageH); }
        var i, lay;
        var base = layoutAt(groups, CFG.baseFraction, pageW, pageH);
        if (!base.fits) {
            for (i = 0; i < CFG.fractions.length; i++) {
                if (CFG.fractions[i] >= CFG.baseFraction) { continue; }
                lay = layoutAt(groups, CFG.fractions[i], pageW, pageH);
                if (lay.fits) { return lay; }
            }
            return lay;
        }
        for (i = 0; i < CFG.fractions.length; i++) {
            if (CFG.fractions[i] <= CFG.baseFraction) { break; }
            lay = layoutAt(groups, CFG.fractions[i], pageW, pageH);
            if (lay.fits && lay.pageCount <= base.pageCount) { return lay; }
        }
        return base;
    }

    // Local positions of every sheet inside a frame (relative to the page,
    // x right / y down). Called after packing.
    function sheetPositions(fr, f) {
        var out = [];
        var y = fr.y + CFG.headerH + CFG.padTop;
        for (var r = 0; r < fr.rows.length; r++) {
            var row = fr.rows[r];
            var x = fr.x + (fr.w - row.w) / 2;
            var badgeSpace = row.hasBadge ? CFG.badgeH + CFG.badgeGap : 0;
            for (var c = 0; c < row.cells.length; c++) {
                var cell = row.cells[c];
                var sw = cell.sheet.w * f;
                var sh = cell.sheet.h * f;
                out.push({
                    sheet: cell.sheet,
                    cx: x + cell.cw / 2,
                    badgeTop: y,
                    labelTop: y + badgeSpace,
                    left: x + (cell.cw - sw) / 2,
                    top: y + badgeSpace + CFG.labelH + CFG.labelGap,
                    w: sw, h: sh
                });
                x += cell.cw + CFG.cellGapX;
            }
            y += row.h + CFG.rowGap;
        }
        return out;
    }

    return { CFG: CFG, chooseLayout: chooseLayout, layoutAt: layoutAt, sheetPositions: sheetPositions };
})();
// ==== END PPL LAYOUT CORE ====

(function () {
    var SCRIPT_NAME = 'ProdProof Layout';
    var TAG = 'PPL';                         // prefix on everything this script creates
    var EXT_RE = /\.(pdf|ai)$/i;
    var SKIP_FOLDER_RE = /^(PRIME|Customer Files|Archive|Old|samples)$/i;
    var FRAME_LAYER = 'Material Frames';
    var ART_LAYER = 'ART';

    if (app.documents.length === 0) {
        alert('Open the ProdProof first, then run ' + SCRIPT_NAME + '.');
        return;
    }
    var doc = app.activeDocument;
    var CFG = PPL.CFG;
    var COVER_MARK = 'following page';       // "Please see proof(s) on following page(s)"
    var oldCoord = app.coordinateSystem;
    app.coordinateSystem = CoordinateSystem.DOCUMENTCOORDINATESYSTEM;

    // ------------------------------------------------------------------
    // Where the new pages go: straight after the last cover page
    // ------------------------------------------------------------------
    function artboardAt(x, y) {
        for (var i = 0; i < doc.artboards.length; i++) {
            var r = doc.artboards[i].artboardRect;
            if (x >= r[0] - 1 && x <= r[2] + 1 && y <= r[1] + 1 && y >= r[3] - 1) { return i; }
        }
        return -1;
    }

    // Index of the last artboard carrying the cover-page line, or -1
    function findLastCoverPage() {
        var last = -1;
        // The proof builder names the cover artboards Cover_1, Cover_2, ...
        for (var a = 0; a < doc.artboards.length; a++) {
            if (/^Cover_\d+$/i.test(doc.artboards[a].name)) { last = a; }
        }
        if (last >= 0) { return last; }
        var frames = doc.textFrames;
        for (var i = 0; i < frames.length; i++) {
            try {
                var tf = frames[i];
                if (String(tf.contents).toLowerCase().indexOf(COVER_MARK) === -1) { continue; }
                var b = tf.geometricBounds;
                var idx = artboardAt((b[0] + b[2]) / 2, (b[1] + b[3]) / 2);
                if (idx > last) { last = idx; }
            } catch (e) {}
        }
        return last;
    }

    // ------------------------------------------------------------------
    // Find the production files the PRIME exporter made
    // ------------------------------------------------------------------
    function docBaseName() {
        var n = doc.name.replace(/\.[^\.]+$/, '');
        return n.replace(/[_\s]*(EmProdProof|ProdProof|EmProof|Proof|PRIME)$/i, '');
    }

    function guessJobFolder() {
        try {
            var here = doc.path;                       // throws when never saved
            if (!here || !here.exists) { return null; }
            if (/^PRIME$/i.test(decodeURI(here.name)) && here.parent) { return here.parent; }
            return here;
        } catch (e) { return null; }
    }

    function scanFolder(folder, depth, out) {
        var list = folder.getFiles();
        for (var i = 0; i < list.length; i++) {
            var f = list[i];
            var nm = decodeURI(f.name);
            if (nm.charAt(0) === '.') { continue; }
            if (f instanceof Folder) {
                if (depth < 3 && !SKIP_FOLDER_RE.test(nm)) { scanFolder(f, depth + 1, out); }
            } else if (EXT_RE.test(nm)) {
                var p = ProdName.parseFileName(nm);
                if (p.type) { out.push({ file: f, name: nm, base: p.base }); }
            }
        }
    }

    function findProductionFiles(folder) {
        var all = [];
        scanFolder(folder, 0, all);
        var want = ProdName.baseKey(docBaseName());
        var mine = [];
        for (var i = 0; i < all.length; i++) {
            if (ProdName.baseKey(all[i].base) === want) { mine.push(all[i]); }
        }
        return { files: mine.length ? mine : all, matchedJob: mine.length > 0, total: all.length };
    }

    // ------------------------------------------------------------------
    // Dialog
    // ------------------------------------------------------------------
    function showDialog(startFolder, coverIndex) {
        var result = null;
        var found = null;
        var dlg = new Window('dialog', SCRIPT_NAME);
        dlg.orientation = 'column';
        dlg.alignChildren = 'fill';

        var p1 = dlg.add('panel', undefined, 'Production files folder (the order folder)');
        p1.orientation = 'row';
        p1.alignChildren = 'center';
        var pathText = p1.add('edittext', undefined, startFolder ? startFolder.fsName : '');
        pathText.characters = 55;
        var browse = p1.add('button', undefined, 'Browse...');
        var status = dlg.add('statictext', undefined, '');
        status.characters = 70;

        var p2 = dlg.add('panel', undefined, 'Options');
        p2.orientation = 'column';
        p2.alignChildren = 'left';
        var g1 = p2.add('group');
        g1.add('statictext', undefined, 'Production files are:');
        var sizeList = g1.add('dropdownlist', undefined,
            ['Auto-detect', '1/10 scale (standard canvas)', 'Full size (Large Canvas)']);
        sizeList.selection = 0;
        p2.add('statictext', undefined, 'The page scale is chosen in the next step, with the page count for each option.');
        var cutBox = p2.add('checkbox', undefined, 'Overlay the CUT file on each PRINT (same as Link Namer)');
        cutBox.value = true;
        var g2 = p2.add('group');
        g2.add('statictext', undefined, 'Add the new pages after artboard:');
        var afterText = g2.add('edittext', undefined,
            String((coverIndex >= 0 ? coverIndex : doc.artboards.getActiveArtboardIndex()) + 1));
        afterText.characters = 4;
        g2.add('statictext', undefined, coverIndex >= 0 ? '(last cover page found)'
                                                         : '(no cover page found - using the active artboard)');

        var btns = dlg.add('group');
        btns.alignment = 'right';
        btns.add('button', undefined, 'Cancel', { name: 'cancel' });
        var ok = btns.add('button', undefined, 'Build pages', { name: 'ok' });

        function refresh() {
            found = null;
            var fld = new Folder(pathText.text);
            if (!pathText.text || !fld.exists) {
                status.text = 'Choose the order folder that holds the material folders.';
                ok.enabled = false;
                return;
            }
            found = findProductionFiles(fld);
            if (!found.files.length) {
                status.text = 'No production files (PRINT / CUT / PnC / CutOnly) found in that folder.';
                ok.enabled = false;
            } else {
                status.text = found.files.length + ' production files found' +
                    (found.matchedJob ? ' for ' + docBaseName() + '.' : ' - NONE match this ProdProof name, using all of them.');
                ok.enabled = true;
            }
        }
        browse.onClick = function () {
            var start = new Folder(pathText.text);
            var pick = (start.exists ? start : Folder.desktop).selectDlg('Select the order folder');
            if (pick) { pathText.text = pick.fsName; refresh(); }
        };
        pathText.onChange = refresh;
        ok.onClick = function () {
            if (!found || !found.files.length) { return; }
            var fileScale = [0, 10, 1][sizeList.selection.index];   // 0 = auto
            var after = parseInt(afterText.text, 10);
            if (isNaN(after) || after < 1 || after > doc.artboards.length) {
                alert('Enter an artboard number from 1 to ' + doc.artboards.length + '.');
                return;
            }
            result = { files: found.files, fileScale: fileScale, cuts: cutBox.value, matchedJob: found.matchedJob,
                       afterIndex: after - 1 };
            dlg.close(1);
        };
        refresh();
        if (dlg.show() !== 1) { return null; }
        return result;
    }

    // ------------------------------------------------------------------
    // Material colours: read back what the Editor sent to this proof. The
    // proof builder fills a path named *_MATERIAL_BG_COLOR and colours the
    // text frame named [*_MATERIAL] on every line page and cover-table row.
    // Returns matKey -> { bg, ink, name }.
    // ------------------------------------------------------------------
    function findBgPath(container, depth) {
        try {
            var items = container.pageItems;
            for (var i = 0; i < items.length; i++) {
                var it = items[i];
                if (it.typename === 'PathItem' && /_MATERIAL_BG_COLOR$/.test(it.name)) { return it; }
            }
            if (depth > 0) {
                for (var g = 0; g < container.groupItems.length; g++) {
                    var hit = findBgPath(container.groupItems[g], depth - 1);
                    if (hit) { return hit; }
                }
            }
        } catch (e) {}
        return null;
    }

    function readMaterialColors() {
        var map = {};
        var frames = doc.textFrames;
        for (var i = 0; i < frames.length; i++) {
            try {
                var tf = frames[i];
                if (!/^\[(LINE1|TABLE|SUB)_MATERIAL\]$/.test(tf.name)) { continue; }
                var mat = String(tf.contents).replace(/^\s+|\s+$/g, '');
                if (!mat) { continue; }
                var key = ProdName.matKey(mat);
                if (map[key]) { continue; }
                // The background path sits beside the text, in the same row / header group
                var holder = tf.parent;
                var bgPath = null;
                for (var up = 0; up < 3 && holder && holder.typename === 'GroupItem' && !bgPath; up++) {
                    bgPath = findBgPath(holder, 2);
                    holder = holder.parent;
                }
                if (!bgPath || !bgPath.filled) { continue; }
                map[key] = { bg: bgPath.fillColor, ink: tf.textRange.characterAttributes.fillColor, name: mat };
            } catch (e) {}
        }
        return map;
    }

    // ------------------------------------------------------------------
    // File scale: standard-canvas PRIMEs export at 1/10, Large Canvas ones
    // at full size. Returns 10 or 1 (how much to multiply a file inch by).
    // ------------------------------------------------------------------
    function hasUserUnit(file) {
        try {
            var f = new File(file.fsName);
            f.encoding = 'BINARY';
            if (!f.open('r')) { return false; }
            var chunk = f.read(400000);
            f.close();
            return chunk.indexOf('/UserUnit') !== -1;
        } catch (e) { return false; }
    }

    function detectFileScale(entries, sizes) {
        var maxShort = 0;
        var maxLong = 0;
        for (var i = 0; i < sizes.length; i++) {
            if (!sizes[i].w) { continue; }
            var a = Math.min(sizes[i].w, sizes[i].h) / 72;
            var b = Math.max(sizes[i].w, sizes[i].h) / 72;
            if (a > maxShort) { maxShort = a; }
            if (b > maxLong) { maxLong = b; }
        }
        // At 1/10 nothing is wider than ~13in or longer than ~100in on the page.
        if (maxShort > 13 || maxLong > 100) { return { scale: 1, why: 'sheet sizes' }; }
        for (var j = 0; j < entries.length && j < 3; j++) {
            if (hasUserUnit(entries[j].file)) { return { scale: 1, why: 'Large Canvas PDF' }; }
        }
        return { scale: 10, why: 'sheet sizes' };
    }

    // ------------------------------------------------------------------
    // Scale picker: every workable scale with the pages it needs
    // ------------------------------------------------------------------
    function pickLayout(groups, pageW, pageH, sheetCount, fileNote) {
        var auto = PPL.chooseLayout(groups, pageW, pageH, 0);
        var options = [];
        var labels = [];
        var autoIdx = 0;
        var fr = CFG.fractions;
        for (var i = 0; i < fr.length; i++) {
            var lay = (fr[i] === auto.fraction) ? auto : PPL.layoutAt(groups, fr[i], pageW, pageH);
            if (!lay.fits) { continue; }
            if (fr[i] === auto.fraction) { autoIdx = options.length; }
            options.push(lay);
            labels.push('1:' + Math.round(PPLModel.FILE_SCALE / fr[i]) + '   -   ' + lay.pageCount +
                        (lay.pageCount === 1 ? ' page' : ' pages') + (fr[i] === auto.fraction ? '   (suggested)' : ''));
        }
        if (!options.length) { return auto; }
        var chosen = null;
        var dlg = new Window('dialog', SCRIPT_NAME + ' - scale');
        dlg.orientation = 'column';
        dlg.alignChildren = 'fill';
        dlg.add('statictext', undefined, sheetCount + ' sheets in ' + groups.length + ' material groups. ' + fileNote);
        dlg.add('statictext', undefined, 'Pick the scale for the production pages:');
        var list = dlg.add('listbox', undefined, labels);
        list.preferredSize = [360, Math.min(260, 22 * labels.length + 10)];
        list.selection = autoIdx;
        var btns = dlg.add('group');
        btns.alignment = 'right';
        btns.add('button', undefined, 'Cancel', { name: 'cancel' });
        var ok = btns.add('button', undefined, 'Build pages', { name: 'ok' });
        ok.onClick = function () {
            if (list.selection) { chosen = options[list.selection.index]; dlg.close(1); }
        };
        if (dlg.show() !== 1) { return null; }
        return chosen;
    }

    // ------------------------------------------------------------------
    // Drawing helpers
    // ------------------------------------------------------------------
    function rgb(r, g, b) {
        var c = new RGBColor();
        c.red = r; c.green = g; c.blue = b;
        return c;
    }

    function fontByNames(names) {
        for (var i = 0; i < names.length; i++) {
            try { return app.textFonts.getByName(names[i]); } catch (e) {}
        }
        return null;
    }
    var FONT_LABEL = fontByNames(['MyriadPro-Black', 'MyriadPro-Bold', 'Arial-BoldMT']);
    var FONT_HEADER = fontByNames(['MyriadPro-BlackSemiCn', 'MyriadPro-Black', 'MyriadPro-Bold', 'Arial-BoldMT']);
    var FONT_BOLDIT = fontByNames(['MyriadPro-BoldIt', 'MyriadPro-Bold', 'Arial-BoldItalicMT']);
    var FONT_NOTE = fontByNames(['MyriadPro-Semibold', 'MyriadPro-Bold', 'Arial-BoldMT']);

    function layerIndex(name) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === name) { return i; }
        }
        return -1;
    }

    var touchedLayers = [];
    function openLayer(layer) {
        touchedLayers.push({ layer: layer, locked: layer.locked, visible: layer.visible });
        layer.locked = false;
        layer.visible = true;
        return layer;
    }
    function restoreLayers() {
        for (var i = 0; i < touchedLayers.length; i++) {
            try {
                touchedLayers[i].layer.locked = touchedLayers[i].locked;
                touchedLayers[i].layer.visible = touchedLayers[i].visible;
            } catch (e) {}
        }
    }

    // Text with its top edge at docTop, centred on docCx
    function addText(container, str, font, size, color, docCx, docTop) {
        var tf = container.textFrames.add();
        tf.contents = str;
        var ca = tf.textRange.characterAttributes;
        if (font) { try { ca.textFont = font; } catch (e) {} }
        ca.size = size;
        ca.fillColor = color;
        try { tf.textRange.paragraphAttributes.justification = Justification.CENTER; } catch (e2) {}
        tf.left = docCx - tf.width / 2;
        tf.top = docTop;
        return tf;
    }

    function addRect(container, docLeft, docTop, w, h, fill, stroke, strokeW) {
        var r = container.pathItems.rectangle(docTop, docLeft, w, h);
        if (fill) { r.filled = true; r.fillColor = fill; } else { r.filled = false; }
        if (stroke) { r.stroked = true; r.strokeColor = stroke; r.strokeWidth = strokeW || 1; } else { r.stroked = false; }
        return r;
    }

    function placeFile(layer, file) {
        var pi = layer.placedItems.add();
        pi.file = file;
        return pi;
    }

    // ------------------------------------------------------------------
    // Artboards
    // ------------------------------------------------------------------
    // Always adds NEW artboards, in page order, straight after afterIndex.
    // Existing pages are never drawn on. On the canvas the new artboards sit
    // in a row below everything else, so they cannot overlap existing art.
    function addArtboards(afterIndex, count, added) {
        var first = doc.artboards[afterIndex].artboardRect;
        var w = first[2] - first[0];
        var h = first[1] - first[3];
        var minBottom = first[3];
        var minLeft = first[0];
        for (var i = 0; i < doc.artboards.length; i++) {
            var r = doc.artboards[i].artboardRect;
            if (r[3] < minBottom) { minBottom = r[3]; }
            if (r[0] < minLeft) { minLeft = r[0]; }
        }
        var rects = [];
        for (var k = 0; k < count; k++) {
            var idx = afterIndex + 1 + k;
            var left = minLeft + k * (w + 40);
            var top = minBottom - 60;
            var nr = [left, top, left + w, top - h];
            try {
                doc.artboards.insert(nr, idx);
            } catch (eIns) {
                doc.artboards.add(nr);                 // lands last; move it in the Artboards panel
                idx = doc.artboards.length - 1;
                added.atEnd = true;
            }
            try { doc.artboards[idx].name = 'Production ' + (k + 1); } catch (e) {}
            added.push(idx + 1);
            rects.push(doc.artboards[idx].artboardRect);
        }
        return rects;
    }

    // ------------------------------------------------------------------
    // Main
    // ------------------------------------------------------------------
    var opts = showDialog(guessJobFolder(), findLastCoverPage());
    if (!opts) { app.coordinateSystem = oldCoord; return; }

    var afterIndex = opts.afterIndex;
    var oldLevel = app.userInteractionLevel;
    app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

    var problems = [];
    var report = null;
    var placed = {};          // path -> PlacedItem (prints)
    try {
        // Layers: art on ART (or the active layer), frames on Material Frames
        // when that layer sits below the art, otherwise at the back of the art layer.
        var artIdx = layerIndex(ART_LAYER);
        var artLayer = openLayer(artIdx >= 0 ? doc.layers[artIdx] : doc.activeLayer);
        artIdx = layerIndex(artLayer.name);
        var frameIdx = layerIndex(FRAME_LAYER);
        var frameLayer = artLayer;
        if (frameIdx > artIdx) { frameLayer = openLayer(doc.layers[frameIdx]); }

        // 1. Place every sheet once to read its real size
        var sizes = [];
        var cutFiles = {};
        var i;
        for (i = 0; i < opts.files.length; i++) {
            var entry = opts.files[i];
            var parsed = ProdName.parseFileName(entry.name);
            if (parsed.type === 'CUT') {
                cutFiles[entry.file.fsName] = entry.file;
                sizes.push({ name: entry.name, path: entry.file.fsName, w: 0, h: 0 });
                continue;
            }
            try {
                var item = placeFile(artLayer, entry.file);
                placed[entry.file.fsName] = item;
                sizes.push({ name: entry.name, path: entry.file.fsName, w: item.width, h: item.height });
            } catch (ePlace) {
                problems.push('Could not place ' + entry.name + ' (' + ePlace + ')');
            }
        }

        // 2. File scale: bring every size to its 1/10 equivalent
        var det = opts.fileScale ? { scale: opts.fileScale, why: 'set in the dialog' } : detectFileScale(opts.files, sizes);
        var kFile = det.scale / PPLModel.FILE_SCALE;      // 1 for 1/10 files, 0.1 for full-size files
        for (i = 0; i < sizes.length; i++) { sizes[i].w *= kFile; sizes[i].h *= kFile; }
        var fileNote = (det.scale === 1 ? 'Files read as FULL SIZE' : 'Files read as 1/10 scale') + ' (' + det.why + ').';

        // 3. Group by material, then let the user pick the scale / page count
        var model = PPLModel.build(sizes);
        if (!model.sheetCount) { throw new Error('None of the files could be placed.'); }
        var ab0 = doc.artboards[afterIndex].artboardRect;
        var pageW = ab0[2] - ab0[0];
        var pageH = ab0[1] - ab0[3];
        app.userInteractionLevel = oldLevel;
        var lay = pickLayout(model.groups, pageW, pageH, model.sheetCount, fileNote);
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;
        if (!lay) { throw new Error('Cancelled - nothing was added.'); }
        var f = lay.fraction;

        // 4. Artboards
        var added = [];
        var rects = addArtboards(afterIndex, lay.pageCount, added);

        // 5. Draw
        var WHITE = rgb(255, 255, 255);
        var BLACK = rgb(0, 0, 0);
        var matColors = readMaterialColors();
        var noColor = [];
        for (i = 0; i < model.groups.length; i++) {
            model.groups[i].colors = matColors[ProdName.matKey(model.groups[i].material)] || null;
            if (!model.groups[i].colors) { noColor.push(model.groups[i].material); }
        }
        var cutMissing = 0;
        var cutMismatch = [];
        for (i = 0; i < lay.frames.length; i++) {
            var fr = lay.frames[i];
            var ab = rects[fr.page];
            var style = fr.group.style;
            var X = function (x) { return ab[0] + x; };
            var Y = function (y) { return ab[1] - y; };

            // Frame: body + header strip + header text, grouped
            var fg = frameLayer.groupItems.add();
            fg.name = TAG + ' frame - ' + fr.group.header;
            var bodyFill = (style === 'blkbk') ? rgb(51, 48, 51) : rgb(227, 227, 227);
            // Header = the material's own colours from the Editor. No match
            // on the proof -> neutral grey (never a guess that could read as
            // a different material).
            var mc = fr.group.colors;
            var headFill = mc ? mc.bg : rgb(242, 242, 242);
            var headInk = mc ? mc.ink : rgb(35, 31, 32);
            var body = addRect(fg, X(fr.x), Y(fr.y + CFG.headerH), fr.w, fr.h - CFG.headerH, bodyFill, null);
            var head = addRect(fg, X(fr.x), Y(fr.y), fr.w, CFG.headerH, headFill, null);
            var ht = addText(fg, fr.group.header + (fr.part > 0 ? ' (cont.)' : ''), FONT_HEADER, 13, headInk,
                             X(fr.x + fr.w / 2), Y(fr.y + 4));
            if (fr.note) {
                addText(fg, fr.note, FONT_NOTE, 9, rgb(220, 30, 36), X(fr.x + fr.w / 2), Y(fr.y + fr.h - CFG.noteH));
            }
            fg.zOrder(ZOrderMethod.SENDTOBACK);

            // Sheets
            var pos = PPL.sheetPositions(fr, f);
            for (var s = 0; s < pos.length; s++) {
                var ps = pos[s];
                var sheet = ps.sheet;
                var print = placed[sheet.path];
                if (!print) { continue; }
                var sg = artLayer.groupItems.add();
                sg.name = TAG + ' ' + sheet.label;

                var back = addRect(sg, X(ps.left), Y(ps.top), ps.w, ps.h, WHITE, null);
                print.resize(f * kFile * 100, f * kFile * 100);
                print.left = X(ps.left);
                print.top = Y(ps.top);
                print.moveToBeginning(sg);

                if (opts.cuts && sheet.cut && cutFiles[sheet.cut.path]) {
                    try {
                        var cut = placeFile(artLayer, cutFiles[sheet.cut.path]);
                        var ratioW = cut.width * kFile / sheet.w;
                        var ratioH = cut.height * kFile / sheet.h;
                        cut.resize(f * kFile * 100, f * kFile * 100);
                        cut.left = X(ps.left);
                        cut.top = Y(ps.top);
                        cut.moveToBeginning(sg);
                        if (Math.abs(ratioW - 1) > 0.02 || Math.abs(ratioH - 1) > 0.02) { cutMismatch.push(sheet.label); }
                    } catch (eCut) {
                        problems.push('Could not place CUT for ' + sheet.label);
                    }
                } else if (opts.cuts && sheet.type === 'PRINT' && sheet.side !== 'BACK') {
                    cutMissing++;
                }

                var ink = (style === 'blkbk') ? WHITE : BLACK;
                var lab = addText(sg, sheet.label, FONT_LABEL, 13, ink, X(ps.cx), Y(ps.labelTop));

                if (sheet.mult > 1) {
                    var bw = 59;
                    var badge = sg.pathItems.roundedRectangle(Y(ps.badgeTop), X(ps.cx - bw / 2), bw, CFG.badgeH, 7, 7);
                    badge.filled = true;
                    badge.fillColor = rgb(179, 222, 194);
                    badge.stroked = true;
                    badge.strokeColor = rgb(0, 102, 46);
                    badge.strokeWidth = 2;
                    var bt = addText(sg, 'x' + sheet.mult, FONT_LABEL, 20, rgb(201, 37, 43), X(ps.cx), Y(ps.badgeTop + 1));
                    try {
                        var xc = bt.textRange.characters[0].characterAttributes;
                        if (FONT_BOLDIT) { xc.textFont = FONT_BOLDIT; }
                        xc.size = 17;
                        xc.fillColor = rgb(0, 101, 46);
                        bt.left = X(ps.cx) - bt.width / 2;
                        bt.top = Y(ps.badgeTop + (CFG.badgeH - bt.height) / 2);
                    } catch (eB) {}
                }
            }
        }

        report = { model: model, lay: lay, added: added, cutMissing: cutMissing, cutMismatch: cutMismatch, fileNote: fileNote, noColor: noColor };
    } catch (err) {
        problems.push('Stopped: ' + err + (err.line ? ' (line ' + err.line + ')' : ''));
        // Do not leave full-size sheets lying around: remove any that were
        // placed but never laid out.
        for (var pth in placed) {
            if (placed.hasOwnProperty(pth)) {
                try { if (placed[pth].parent.typename !== 'GroupItem') { placed[pth].remove(); } } catch (eRm) {}
            }
        }
    }

    restoreLayers();
    try { if (report && report.added.length) { doc.artboards.setActiveArtboardIndex(report.added[0] - 1); } } catch (eAb) {}
    doc.selection = null;
    app.coordinateSystem = oldCoord;
    app.userInteractionLevel = oldLevel;

    // ------------------------------------------------------------------
    // Summary: what was decided, what still needs a person
    // ------------------------------------------------------------------
    var msg = [];
    if (report) {
        var denom = Math.round(PPLModel.FILE_SCALE / report.lay.fraction);
        msg.push(report.model.sheetCount + ' sheets in ' + report.model.groups.length + ' material groups, ' +
                 report.lay.pageCount + ' new page(s) added after artboard ' + (afterIndex + 1) + '.');
        msg.push('Scale 1:' + denom + '. ' + report.fileNote);
        var sf = 1;
        try { sf = doc.scaleFactor || 1; } catch (eSf) {}
        if (sf !== 1) { msg.push('WARNING: this ProdProof is itself a Large Canvas document (x' + sf + '); sizes and spacing may be off.'); }
        if (!report.lay.fits) { msg.push('WARNING: at this scale something is larger than the page. Pick a smaller scale and run again.'); }
        msg.push('New artboard(s): ' + report.added.join(', ') + '. On the canvas they sit below the existing artboards' +
                 (report.added.atEnd ? '; they could not be inserted in order, so drag them up in the Artboards panel' : '') + '.');
        if (!opts.matchedJob) { msg.push('WARNING: no file names matched this ProdProof, so every production file in the folder was used.'); }
        if (report.model.multiples.length) {
            msg.push('');
            msg.push('MULTIPLES (badges added):');
            for (var m = 0; m < report.model.multiples.length; m++) { msg.push('  ' + report.model.multiples[m]); }
        }
        var w = report.model.warnings;
        if (w.length || report.cutMissing || report.cutMismatch.length || report.noColor.length) {
            msg.push('');
            msg.push('CHECK:');
            for (var q = 0; q < w.length; q++) { msg.push('  ' + w[q]); }
            if (report.cutMissing) { msg.push('  ' + report.cutMissing + ' PRINT sheet(s) have no CUT file'); }
            if (report.cutMismatch.length) { msg.push('  CUT size differs from PRINT: ' + report.cutMismatch.join(', ')); }
            if (report.noColor.length) { msg.push('  No material colour on this proof (grey header used): ' + report.noColor.join(', ')); }
        }
        msg.push('');
        msg.push('STILL BY HAND: VHB / Black Back tags, color match swatches, reflective / laminate flags, external production table, job notes.');
    }
    if (problems.length) {
        msg.push('');
        msg.push('PROBLEMS:');
        for (var z = 0; z < problems.length; z++) { msg.push('  ' + problems[z]); }
    }
    alert(SCRIPT_NAME + '\n\n' + msg.join('\n'));
})();
// end of script
