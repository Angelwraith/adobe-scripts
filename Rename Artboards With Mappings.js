/*
@METADATA
{
  "name": "Rename Artboards With Mappings",
  "description": "Opens a dialog with a multi-select list of all artboards. Add one or more mapping rows below; each row captures a set of selected artboards, a name (free-text), and a starting _Pt number. On Rename, each mapping renames its captured artboards to 'Name_PtN' incrementing N from the chosen start.",
  "version": "1.0",
  "target": "illustrator",
  "tags": ["artboard", "rename", "mapping", "ui"]
}
@END_METADATA
*/

// Future: when a CEP bridge function in the launcher can deliver a live
// list of company materials, populate MATERIAL_PRESETS below and the row
// UI will switch the name input from a text field to a dropdown+text combo.
var MATERIAL_PRESETS = [];

(function () {
    'use strict';

    if (app.documents.length === 0) {
        alert("Please open a document first.");
        return;
    }

    var doc = app.activeDocument;
    var boards = doc.artboards;

    if (boards.length === 0) {
        alert("This document has no artboards.");
        return;
    }

    var mappings = showRenamerDialog();
    if (mappings === null) return; // cancelled

    var summary = applyMappings(mappings);
    alert(summary);

    /* -------------------------------------------------------------- */

    function showRenamerDialog() {
        var dialog = new Window("dialog", "Rename Artboards With Mappings");
        dialog.orientation = "column";
        dialog.alignChildren = "fill";
        dialog.spacing = 12;
        dialog.margins = 20;
        dialog.preferredSize.width = 620;

        var header = dialog.add("statictext", undefined,
            "Select artboards above, then add mapping rows below to rename them.");
        header.graphics.font = ScriptUI.newFont("dialog", "bold", 12);

        // ---- Artboard list ----
        var listHeight = Math.min(boards.length * 20 + 10, 350);
        var list = dialog.add("listbox", undefined, [], {multiselect: true});
        list.preferredSize.width = 580;
        list.preferredSize.height = listHeight;

        for (var i = 0; i < boards.length; i++) {
            list.add("item", (i + 1) + ":  " + boards[i].name);
        }

        // Default: all selected
        var allIndices = [];
        for (var iAll = 0; iAll < boards.length; iAll++) allIndices.push(iAll);
        list.selection = allIndices;

        var help = dialog.add("statictext", undefined,
            "Ctrl/Cmd+Click toggles individual items. Shift+Click selects a range.");
        help.graphics.font = ScriptUI.newFont("dialog", "italic", 10);

        var selBtns = dialog.add("group");
        selBtns.alignment = "center";
        selBtns.spacing = 10;
        var selAllBtn = selBtns.add("button", undefined, "Select All");
        selAllBtn.preferredSize.width = 100;
        var deselAllBtn = selBtns.add("button", undefined, "Deselect All");
        deselAllBtn.preferredSize.width = 100;
        selAllBtn.onClick = function () {
            var all = [];
            for (var i = 0; i < boards.length; i++) all.push(i);
            list.selection = all;
        };
        deselAllBtn.onClick = function () { list.selection = null; };

        // ---- Divider ----
        var divider = dialog.add("panel");
        divider.preferredSize.height = 2;
        divider.alignment = "fill";

        // ---- Mappings section ----
        var mapHeader = dialog.add("statictext", undefined, "Mappings:");
        mapHeader.graphics.font = ScriptUI.newFont("dialog", "bold", 11);

        var rowsContainer = dialog.add("group");
        rowsContainer.orientation = "column";
        rowsContainer.alignChildren = "fill";
        rowsContainer.spacing = 6;

        var rows = []; // rowData objects in display order

        function addRow() {
            var rowData = {
                indices: [],
                nameField: null,
                startField: null,
                infoText: null,
                group: null
            };

            var rowGroup = rowsContainer.add("group");
            rowGroup.orientation = "row";
            rowGroup.alignChildren = "center";
            rowGroup.spacing = 6;
            rowData.group = rowGroup;

            var captureBtn = rowGroup.add("button", undefined, "Set From Selection");
            captureBtn.preferredSize.width = 130;

            var infoText = rowGroup.add("statictext", undefined, "(no boards)");
            infoText.preferredSize.width = 110;
            rowData.infoText = infoText;

            rowGroup.add("statictext", undefined, "Name:");
            var nameField = rowGroup.add("edittext", undefined, "");
            nameField.preferredSize.width = 140;
            rowData.nameField = nameField;

            rowGroup.add("statictext", undefined, "Start _Pt:");
            var startField = rowGroup.add("edittext", undefined, "1");
            startField.preferredSize.width = 50;
            rowData.startField = startField;

            var removeBtn = rowGroup.add("button", undefined, "X");
            removeBtn.preferredSize.width = 30;

            captureBtn.onClick = function () {
                var sel = getListSelection(list);
                rowData.indices = sel.slice();
                infoText.text = describeIndices(rowData.indices);
                dialog.layout.layout(true);
            };

            removeBtn.onClick = function () {
                try { rowsContainer.remove(rowGroup); } catch (e) {}
                for (var r = 0; r < rows.length; r++) {
                    if (rows[r] === rowData) { rows.splice(r, 1); break; }
                }
                dialog.layout.layout(true);
            };

            rows.push(rowData);

            // Auto-capture whatever is selected at the moment of adding so a
            // typical session is "select boards, add row, type name, repeat".
            var currentSel = getListSelection(list);
            if (currentSel.length > 0) {
                rowData.indices = currentSel.slice();
                infoText.text = describeIndices(rowData.indices);
            }

            dialog.layout.layout(true);
            return rowData;
        }

        // ---- Bottom buttons ----
        var bottomGroup = dialog.add("group");
        bottomGroup.alignment = "center";
        bottomGroup.spacing = 10;

        var addBtn = bottomGroup.add("button", undefined, "+ Add Mapping Row");
        addBtn.preferredSize.width = 160;
        var okBtn = bottomGroup.add("button", undefined, "Rename");
        okBtn.preferredSize.width = 100;
        var cancelBtn = bottomGroup.add("button", undefined, "Cancel");
        cancelBtn.preferredSize.width = 80;

        addBtn.onClick = function () { addRow(); };

        var resultMappings = null;

        okBtn.onClick = function () {
            var result = [];
            var problems = [];
            for (var r = 0; r < rows.length; r++) {
                var row = rows[r];
                var name = (row.nameField.text || "").replace(/^\s+|\s+$/g, '');
                var startStr = (row.startField.text || "").replace(/^\s+|\s+$/g, '');
                var start = parseInt(startStr, 10);
                if (isNaN(start)) start = 1;

                if (row.indices.length === 0 && name === "") {
                    continue; // entirely empty row, just ignore
                }
                if (row.indices.length === 0) {
                    problems.push("Row " + (r + 1) + ": no artboards captured. Click 'Set From Selection'.");
                    continue;
                }
                if (name === "") {
                    problems.push("Row " + (r + 1) + ": name is empty.");
                    continue;
                }

                var sorted = row.indices.slice().sort(function (a, b) { return a - b; });
                result.push({ indices: sorted, name: name, start: start });
            }

            if (problems.length > 0) {
                alert("Fix these before renaming:\n\n" + problems.join("\n"));
                return;
            }
            if (result.length === 0) {
                alert("No valid mappings to apply. Add a mapping row, capture a selection, and type a name.");
                return;
            }

            resultMappings = result;
            dialog.close();
        };

        cancelBtn.onClick = function () {
            resultMappings = null;
            dialog.close();
        };

        // Start with one row pre-populated
        addRow();

        dialog.show();
        return resultMappings;
    }

    /**
     * Pull selected indices out of a multi-select listbox, normalised
     * to a sorted array of integers.
     */
    function getListSelection(list) {
        var indices = [];
        if (list.selection) {
            if (list.selection.length !== undefined) {
                for (var i = 0; i < list.selection.length; i++) {
                    indices.push(list.selection[i].index);
                }
            } else {
                indices.push(list.selection.index);
            }
        }
        indices.sort(function (a, b) { return a - b; });
        return indices;
    }

    /**
     * Compact summary of an index list: "1-3, 5, 7-9" if it fits,
     * otherwise "N boards".
     */
    function describeIndices(indices) {
        if (indices.length === 0) return "(no boards)";

        var ranges = [];
        var start = indices[0];
        var prev = start;
        for (var i = 1; i < indices.length; i++) {
            if (indices[i] === prev + 1) {
                prev = indices[i];
            } else {
                ranges.push(start === prev ? String(start + 1) : (start + 1) + "-" + (prev + 1));
                start = prev = indices[i];
            }
        }
        ranges.push(start === prev ? String(start + 1) : (start + 1) + "-" + (prev + 1));

        var rangeStr = ranges.join(", ");
        if (rangeStr.length > 22) {
            return indices.length + " boards";
        }
        return rangeStr;
    }

    /**
     * Apply the captured mappings. If a board appears in more than one
     * mapping, the later one wins and we report the conflict.
     */
    function applyMappings(mappings) {
        var owner = {}; // index -> mapping number that last claimed it
        var collisions = [];
        var renamed = 0;
        var examples = [];

        for (var m = 0; m < mappings.length; m++) {
            var mapping = mappings[m];
            for (var i = 0; i < mapping.indices.length; i++) {
                var idx = mapping.indices[i];
                if (owner[idx] !== undefined) {
                    collisions.push("Board " + (idx + 1) +
                                    " was in mapping #" + (owner[idx] + 1) +
                                    " and again in #" + (m + 1));
                }
                owner[idx] = m;
            }
        }

        for (var mm = 0; mm < mappings.length; mm++) {
            var map = mappings[mm];
            var ptCounter = 0;
            for (var ii = 0; ii < map.indices.length; ii++) {
                var index = map.indices[ii];
                if (owner[index] !== mm) {
                    // Claimed by a later mapping. Skip here so the later
                    // mapping's numbering doesn't get a hole.
                    continue;
                }
                var ptNum = map.start + ptCounter;
                ptCounter++;
                var newName = map.name + "_Pt" + ptNum;
                var oldName = boards[index].name;
                try {
                    boards[index].name = newName;
                    renamed++;
                    if (examples.length < 8) {
                        examples.push("Board " + (index + 1) + ": " + oldName + "  ->  " + newName);
                    }
                } catch (e) {
                    // ignore, count as not renamed
                }
            }
        }

        var msg = "Rename Artboards complete.\n" +
                  "------------------------------\n" +
                  "Mappings applied: " + mappings.length + "\n" +
                  "Artboards renamed: " + renamed;

        if (collisions.length > 0) {
            msg += "\n\nConflicts (later mapping won): " + collisions.length;
            for (var c = 0; c < Math.min(collisions.length, 3); c++) {
                msg += "\n  - " + collisions[c];
            }
            if (collisions.length > 3) {
                msg += "\n  ... and " + (collisions.length - 3) + " more";
            }
        }

        if (examples.length > 0) {
            msg += "\n\nExamples:";
            for (var e = 0; e < examples.length; e++) {
                msg += "\n  " + examples[e];
            }
        }

        return msg;
    }
})();
