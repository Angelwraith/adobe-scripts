/*
@METADATA
{
  "name": "Sort Linked Files By Pt Number",
  "description": "Sort selected linked files left-to-right by the Pt# in their filename (Pt1, Pt2, Pt3...) with 0.25\" spacing between them. Tops align to the topmost item in the current selection; the leftmost X of the selection is the starting position.",
  "version": "1.0",
  "target": "illustrator",
  "tags": ["sort", "linked", "files", "layout", "processor"]
}
@END_METADATA
*/

(function () {
    'use strict';

    // Horizontal spacing between sorted items.
    var SPACING_INCHES = 0.25;
    var SPACING_PTS    = SPACING_INCHES * 72;

    // Matches a Pt followed by digits anywhere in the name (Pt1, Pt07, Pt123).
    // Case-insensitive so "pt1" or "PT1" also work.
    var PT_REGEX = /Pt(\d+)/i;

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

    // Collect placed items that have a Pt# in their filename.
    var items = [];
    var skippedNonPlaced = 0;
    var skippedNoPt = 0;
    var skippedNoPtNames = [];

    for (var i = 0; i < sel.length; i++) {
        var it = sel[i];
        if (it.typename !== "PlacedItem") {
            skippedNonPlaced++;
            continue;
        }

        // Prefer the linked file's name on disk; fall back to the item's
        // own name in the layers panel if there's no linked file.
        var name = "";
        try {
            if (it.file) name = it.file.name;
        } catch (e) {}
        if (!name) {
            try { name = it.name; } catch (e) {}
        }

        var match = name.match(PT_REGEX);
        if (!match) {
            skippedNoPt++;
            if (skippedNoPtNames.length < 5) skippedNoPtNames.push(name || "(unnamed)");
            continue;
        }

        items.push({
            item: it,
            ptNumber: parseInt(match[1], 10),
            name: name
        });
    }

    if (items.length < 2) {
        alert("Need at least 2 linked files with 'Pt#' in the filename.\n" +
              "Found: " + items.length + " usable file(s).");
        return;
    }

    // Sort ascending by Pt number.
    items.sort(function (a, b) { return a.ptNumber - b.ptNumber; });

    // Reference position: leftmost X across the whole selection becomes
    // the starting X; topmost Y across the whole selection becomes the
    // baseline that every sorted item's top is aligned to.
    var startX  =  Infinity;
    var topY    = -Infinity;
    for (var k = 0; k < items.length; k++) {
        var b = items[k].item.geometricBounds; // [L, T, R, B]
        if (b[0] < startX) startX = b[0];
        if (b[1] > topY)   topY   = b[1];
    }

    // Place items left-to-right at the baseline, separated by SPACING_PTS.
    var cursorX = startX;
    for (var j = 0; j < items.length; j++) {
        var pi = items[j].item;
        var pb = pi.geometricBounds;
        var width  = pb[2] - pb[0];
        var curLeft = pb[0];
        var curTop  = pb[1];

        var dx = cursorX - curLeft;
        var dy = topY    - curTop;

        // translate(dx, dy) — dy positive moves up in Illustrator's coord system.
        pi.translate(dx, dy);

        cursorX += width + SPACING_PTS;
    }

    app.redraw();

    var msg = "Sort complete.\n" +
              "------------------------------\n" +
              "Sorted: " + items.length + " linked file(s) by Pt#\n" +
              "Spacing: " + SPACING_INCHES + '" between items\n' +
              "Tops aligned to the topmost item in the selection.";

    if (skippedNonPlaced > 0) {
        msg += "\n\nSkipped (not a linked/placed file): " + skippedNonPlaced;
    }
    if (skippedNoPt > 0) {
        msg += "\n\nSkipped (no Pt# in filename): " + skippedNoPt;
        for (var n = 0; n < skippedNoPtNames.length; n++) {
            msg += "\n  - " + skippedNoPtNames[n];
        }
        if (skippedNoPt > skippedNoPtNames.length) {
            msg += "\n  ...and " + (skippedNoPt - skippedNoPtNames.length) + " more";
        }
    }

    alert(msg);
})();
