/*
@METADATA
{
  "name": "Clip Into Each Mask",
  "description": "Select two groups: the top (front-most) group holds mask shapes, the bottom (back) group holds the content. The content group is duplicated once per mask shape and each duplicate is clipped into one of the shapes from the top group.",
  "version": "1.7",
  "target": "illustrator",
  "tags": ["clip", "mask", "duplicate", "group", "processor"]
}
@END_METADATA
*/

(function () {
    'use strict';

    // Flip to true to step through each loop step with an alert if you
    // ever need to diagnose another failure.
    var DEBUG = false;

    function dbg(msg) {
        if (DEBUG) alert("[trace] " + msg);
    }

    if (app.documents.length === 0) {
        alert("Please open a document first.");
        return;
    }

    var doc = app.activeDocument;
    var sel = doc.selection;

    if (!sel || sel.length !== 2 ||
        sel[0].typename !== "GroupItem" ||
        sel[1].typename !== "GroupItem") {
        alert("Select exactly two groups:\n\n" +
              "  - Top group: holds the mask shapes\n" +
              "  - Bottom group: holds the content to be clipped\n\n" +
              "'Top' and 'bottom' are read from stacking order (front = top).");
        return;
    }

    // Both groups must share a parent for index-based z-order to be
    // meaningful and for the new clip groups to sit beside them.
    var parent = sel[0].parent;
    if (sel[1].parent !== parent) {
        alert("The two selected groups must live in the same parent\n" +
              "(same layer or same enclosing group).");
        return;
    }

    // Decide which selected group is on top vs bottom by index in the
    // shared parent's pageItems list. Lower index = closer to the front.
    //
    // We use this instead of zOrderPosition because that property has
    // been observed to throw 'Internal error' on recent Illustrator
    // builds (2025/2026). Index lookup always works.
    var idx0 = indexInParent(sel[0], parent);
    var idx1 = indexInParent(sel[1], parent);

    var topGroup, bottomGroup;
    if (idx0 < idx1) {
        topGroup    = sel[0];
        bottomGroup = sel[1];
    } else {
        topGroup    = sel[1];
        bottomGroup = sel[0];
    }

    // Collect references to every direct child of the top group up front.
    var maskShapes = [];
    for (var i = 0; i < topGroup.pageItems.length; i++) {
        maskShapes.push(topGroup.pageItems[i]);
    }

    if (maskShapes.length === 0) {
        alert("The top group has no items to use as masks.");
        return;
    }

    var processed = 0;
    var errors = 0;
    var errorMessages = [];
    var createdClipGroups = [];

    for (var k = 0; k < maskShapes.length; k++) {
        var stepLabel = "init";
        try {
            var maskShape = maskShapes[k];

            stepLabel = "duplicate bottom group";
            dbg("Mask " + (k + 1) + ": " + stepLabel);
            var contentCopy = bottomGroup.duplicate();

            stepLabel = "create clip group";
            dbg("Mask " + (k + 1) + ": " + stepLabel);
            var clipGroup = parent.groupItems.add();

            stepLabel = "move content into clip group";
            dbg("Mask " + (k + 1) + ": " + stepLabel);
            contentCopy.moveToEnd(clipGroup);

            stepLabel = "move mask into clip group";
            dbg("Mask " + (k + 1) + ": " + stepLabel);
            maskShape.moveToBeginning(clipGroup);

            stepLabel = "apply clipping flag";
            dbg("Mask " + (k + 1) + ": " + stepLabel);
            applyClipping(maskShape);

            // Mark the group itself as a clip group. Setting clipping on
            // the topmost path is not enough on Illustrator 2026 - the
            // group needs to be explicitly flagged or it just renders as
            // a normal group containing the mask shape and the content.
            stepLabel = "set clipped on group";
            dbg("Mask " + (k + 1) + ": " + stepLabel);
            try { clipGroup.clipped = true; } catch (eClipped) {}

            createdClipGroups.push(clipGroup);
            processed++;
        } catch (e) {
            errors++;
            if (errorMessages.length < 5) {
                errorMessages.push("Mask " + (k + 1) + " [" + stepLabel + "]: " + e.message);
            }
        }
    }

    // Clean up the original wrapper groups.
    try {
        if (topGroup.pageItems.length === 0) {
            topGroup.remove();
        }
    } catch (eTop) {}
    try {
        bottomGroup.remove();
    } catch (eBot) {}

    // Select the new clip groups so the user can keep working with them.
    doc.selection = null;
    for (var s = 0; s < createdClipGroups.length; s++) {
        try { createdClipGroups[s].selected = true; } catch (eSel) {}
    }

    app.redraw();

    var msg = "Clip Into Each Mask complete.\n" +
              "------------------------------\n" +
              "Mask shapes: " + maskShapes.length + "\n" +
              "Clip groups created: " + processed;

    if (errors > 0) {
        msg += "\nErrors: " + errors;
        for (var em = 0; em < errorMessages.length; em++) {
            msg += "\n  - " + errorMessages[em];
        }
    }

    alert(msg);

    /* -------------------------------------------------------------- */

    /**
     * Find the index of an item inside a container's pageItems list.
     * Used as a safe substitute for zOrderPosition, which can throw
     * 'Internal error' on some Illustrator builds. Returns -1 if not
     * found (shouldn't happen for items that are selected inside the
     * container, but defensive anyway).
     */
    function indexInParent(item, container) {
        var items = container.pageItems;
        for (var i = 0; i < items.length; i++) {
            if (items[i] === item) return i;
        }
        return -1;
    }

    /**
     * Flag the topmost item of a clip group as the clipping path.
     * Handles PathItem and CompoundPathItem; recurses into nested
     * GroupItems if the user pre-grouped a complex mask.
     */
    function applyClipping(item) {
        if (!item) return;
        if (item.typename === "PathItem") {
            try { item.clipping = true; } catch (e) {}
            return;
        }
        if (item.typename === "CompoundPathItem") {
            try {
                var subs = item.pathItems;
                for (var p = 0; p < subs.length; p++) {
                    try { subs[p].clipping = true; } catch (e2) {}
                }
            } catch (e3) {}
            return;
        }
        if (item.typename === "GroupItem") {
            var kids = item.pageItems;
            for (var g = 0; g < kids.length; g++) {
                applyClipping(kids[g]);
            }
        }
    }
})();
