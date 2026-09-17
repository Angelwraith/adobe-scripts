/*
@METADATA
{
  "name": "Pathfinder Per Group",
  "description": "Runs any Pathfinder operation (Unite, Minus Front, Intersect, Exclude, Divide, Trim, Merge, Crop, Outline, Minus Back) on each selected group individually. Each group's contents get combined into its own result instead of every group's contents being merged together (which is what happens if you run Pathfinder on the full selection at once).",
  "version": "2.0",
  "target": "illustrator",
  "tags": ["pathfinder", "unite", "exclude", "divide", "trim", "merge", "crop", "outline", "processor", "group"]
}
@END_METADATA
*/

(function () {
    'use strict';

    if (app.documents.length === 0) {
        alert("Please open a document first.");
        return;
    }

    var doc = app.activeDocument;
    var sel = doc.selection;

    if (!sel || sel.length === 0) {
        alert("Select one or more groups to process.");
        return;
    }

    // Collect group references from the selection.
    // (We grab references up-front because the selection will change
    // as we process each group.)
    var groups = [];
    for (var i = 0; i < sel.length; i++) {
        if (sel[i].typename === "GroupItem") {
            groups.push(sel[i]);
        }
    }

    if (groups.length === 0) {
        alert("No groups found in the selection.\n" +
              "Select the groups you want a Pathfinder operation applied to.");
        return;
    }

    // ---------------------------------------------------------------
    // Pathfinder operations, laid out to mirror the Pathfinder panel.
    // Each entry: display label + the menu command for the live effect.
    // ---------------------------------------------------------------
    var SHAPE_MODES = [
        { key: "unite",      label: "Unite",       cmd: "Live Pathfinder Add" },
        { key: "minusFront", label: "Minus Front", cmd: "Live Pathfinder Subtract" },
        { key: "intersect",  label: "Intersect",   cmd: "Live Pathfinder Intersect" },
        { key: "exclude",    label: "Exclude",     cmd: "Live Pathfinder Exclude" }
    ];

    var PATHFINDERS = [
        { key: "divide",    label: "Divide",     cmd: "Live Pathfinder Divide" },
        { key: "trim",      label: "Trim",       cmd: "Live Pathfinder Trim" },
        { key: "merge",     label: "Merge",      cmd: "Live Pathfinder Merge" },
        { key: "crop",      label: "Crop",       cmd: "Live Pathfinder Crop" },
        { key: "outline",   label: "Outline",    cmd: "Live Pathfinder Outline" },
        { key: "minusBack", label: "Minus Back", cmd: "Live Pathfinder Minus Back" }
    ];

    var ALL_OPS = SHAPE_MODES.concat(PATHFINDERS);

    // Remember the last-used operation between runs (default: Exclude,
    // which matches this script's original behavior).
    var PREF_KEY = "PathfinderPerGroup_lastOp";
    var lastKey = "exclude";
    try {
        var stored = app.preferences.getStringPreference(PREF_KEY);
        if (stored) { lastKey = stored; }
    } catch (e) { /* no stored preference yet */ }

    // ---------------------------------------------------------------
    // Dialog: pick the operation to run on each group.
    // ---------------------------------------------------------------
    var chosen = null;

    var dlg = new Window("dialog", "Pathfinder Per Group");
    dlg.orientation = "column";
    dlg.alignChildren = "fill";

    var info = dlg.add("statictext", undefined,
        "Runs the chosen operation on each of the " + groups.length +
        " selected group(s) individually.", { multiline: true });
    info.preferredSize.width = 260;

    var radios = [];

    function addOpPanel(title, ops) {
        var pnl = dlg.add("panel", undefined, title);
        pnl.orientation = "column";
        pnl.alignChildren = "left";
        pnl.margins = [15, 15, 15, 10];
        for (var j = 0; j < ops.length; j++) {
            var rb = pnl.add("radiobutton", undefined, ops[j].label);
            rb.opKey = ops[j].key;
            rb.value = (ops[j].key === lastKey);
            radios.push(rb);
        }
    }

    addOpPanel("Shape Modes:", SHAPE_MODES);
    addOpPanel("Pathfinders:", PATHFINDERS);

    // Make sure exactly one radio is selected even if the stored
    // preference key was invalid.
    var anyChecked = false;
    for (var r = 0; r < radios.length; r++) {
        if (radios[r].value) { anyChecked = true; break; }
    }
    if (!anyChecked) { radios[3].value = true; } // Exclude

    // Radio buttons in separate panels are separate groups in ScriptUI,
    // so uncheck the others manually when one is clicked.
    function onRadioClick() {
        for (var q = 0; q < radios.length; q++) {
            if (radios[q] !== this) { radios[q].value = false; }
        }
    }
    for (var p = 0; p < radios.length; p++) {
        radios[p].onClick = onRadioClick;
    }

    var btnRow = dlg.add("group");
    btnRow.alignment = "right";
    var cancelBtn = btnRow.add("button", undefined, "Cancel", { name: "cancel" });
    var okBtn = btnRow.add("button", undefined, "OK", { name: "ok" });

    okBtn.onClick = function () {
        for (var q = 0; q < radios.length; q++) {
            if (radios[q].value) {
                for (var a = 0; a < ALL_OPS.length; a++) {
                    if (ALL_OPS[a].key === radios[q].opKey) {
                        chosen = ALL_OPS[a];
                        break;
                    }
                }
                break;
            }
        }
        dlg.close(1);
    };

    if (dlg.show() !== 1 || !chosen) {
        return; // cancelled
    }

    try {
        app.preferences.setStringPreference(PREF_KEY, chosen.key);
    } catch (e) { /* preferences unavailable; not critical */ }

    // ---------------------------------------------------------------
    // Apply the chosen operation to each group individually.
    // ---------------------------------------------------------------
    var processed = 0;
    var errors = 0;
    var errorMessages = [];

    for (var g = 0; g < groups.length; g++) {
        try {
            // Isolate this group as the selection
            doc.selection = null;
            groups[g].selected = true;

            // Apply the Pathfinder operation as a live effect on the
            // group, then expand the appearance to bake it into real
            // geometry. This is the scripting equivalent of clicking
            // the corresponding button in the Pathfinder palette for
            // just this group.
            app.executeMenuCommand(chosen.cmd);
            app.executeMenuCommand('expandStyle');

            processed++;
        } catch (e) {
            errors++;
            if (errorMessages.length < 3) {
                errorMessages.push(e.message);
            }
        }
    }

    // Re-select whatever survived so the user can keep working with them.
    doc.selection = null;
    for (var k = 0; k < groups.length; k++) {
        try {
            groups[k].selected = true;
        } catch (e) {
            // Group reference may have been replaced by the expand step.
        }
    }

    app.redraw();

    var msg = "Pathfinder " + chosen.label + " per Group complete.\n" +
              "------------------------------\n" +
              "Groups processed: " + processed + " of " + groups.length;

    if (errors > 0) {
        msg += "\nErrors: " + errors;
        for (var m = 0; m < errorMessages.length; m++) {
            msg += "\n  - " + errorMessages[m];
        }
    }

    alert(msg);
})();
