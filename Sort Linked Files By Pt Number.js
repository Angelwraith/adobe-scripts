/*
@METADATA
{
  "name": "Sort Linked Files By Pt Number",
  "description": "Sort selected linked files. If the files are wrap sides (driver/front/passenger/rear/top in the name, matching Wrap Processor's naming), they're grouped by side and stacked top-to-bottom in the order Driver, Front, Passenger, Rear, Top (missing sides omitted), each row sorted left-to-right by panel number. Otherwise falls back to the original Pt# left-to-right sort. 0.25\" spacing; the leftmost/topmost of the selection is the starting position.",
  "version": "1.2",
  "target": "illustrator",
  "tags": ["sort", "linked", "files", "layout", "wrap", "processor"]
}
@END_METADATA
*/

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
    var wrapCount = 0;
    for (var p = 0; p < placed.length; p++) {
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
            var match = placed[i].name.match(PT_REGEX);
            if (!match) {
                skippedNoPt++;
                if (skippedNoPtNames.length < 5) skippedNoPtNames.push(placed[i].name || "(unnamed)");
                continue;
            }
            items.push({
                item: placed[i].item,
                ptNumber: parseInt(match[1], 10),
                name: placed[i].name
            });
        }

        if (items.length < 2) {
            alert("Need at least 2 linked files with 'Pt#' in the filename.\n" +
                  "Found: " + items.length + " usable file(s).");
            return;
        }

        items.sort(function (a, b) { return a.ptNumber - b.ptNumber; });

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
