/*
@METADATA
{
  "name": "RGB to CMYK Batch",
  "description": "Convert and save selected open documents from RGB to CMYK",
  "version": "1.0",
  "target": "illustrator",
  "tags": ["color", "cmyk", "rgb", "batch", "convert"]
}
@END_METADATA
*/

#target illustrator

(function() {
    if (app.documents.length === 0) {
        alert("No documents are open.");
        return;
    }

    // Snapshot all open documents up front (indexes can shift if docs are closed mid-run)
    var docs = [];
    for (var i = 0; i < app.documents.length; i++) {
        docs.push(app.documents[i]);
    }

    // ---------- Dialog ----------
    var dialog = new Window("dialog", "RGB to CMYK Batch Converter");
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.spacing = 10;
    dialog.margins = 16;

    var header = dialog.add("statictext", undefined, "Select documents to convert to CMYK:");
    header.graphics.font = ScriptUI.newFont("Arial", "BOLD", 12);

    // Scrollable panel of checkboxes
    var panel = dialog.add("panel");
    panel.orientation = "column";
    panel.alignChildren = "left";
    panel.margins = 10;
    panel.spacing = 4;
    panel.preferredSize.width = 520;

    var checkboxes = [];
    for (var i = 0; i < docs.length; i++) {
        var doc = docs[i];
        var colorMode = "Unknown";
        try {
            colorMode = (doc.documentColorSpace === DocumentColorSpace.RGB) ? "RGB" : "CMYK";
        } catch (e) {}

        var hasPath = false;
        try { hasPath = !!(doc.path && doc.path.fsName); } catch (e) {}

        var label = doc.name + "   [" + colorMode + "]";
        if (!hasPath) label += "  (unsaved \u2014 will not be saved)";

        var cb = panel.add("checkbox", undefined, label);
        cb._doc = doc;
        cb._colorMode = colorMode;
        cb._hasPath = hasPath;

        // Pre-check RGB documents that have a save path
        if (colorMode === "RGB" && hasPath) cb.value = true;

        checkboxes.push(cb);
    }

    // Quick selection buttons
    var selectGroup = dialog.add("group");
    selectGroup.alignment = "left";
    selectGroup.spacing = 6;
    var selectAllBtn = selectGroup.add("button", undefined, "All");
    var selectNoneBtn = selectGroup.add("button", undefined, "None");
    var selectRGBBtn = selectGroup.add("button", undefined, "RGB Only");

    selectAllBtn.onClick = function() {
        for (var i = 0; i < checkboxes.length; i++) checkboxes[i].value = true;
    };
    selectNoneBtn.onClick = function() {
        for (var i = 0; i < checkboxes.length; i++) checkboxes[i].value = false;
    };
    selectRGBBtn.onClick = function() {
        for (var i = 0; i < checkboxes.length; i++) {
            checkboxes[i].value = (checkboxes[i]._colorMode === "RGB");
        }
    };

    // Save option
    var saveGroup = dialog.add("group");
    saveGroup.alignment = "left";
    var saveCheckbox = saveGroup.add("checkbox", undefined, "Save each document after converting");
    saveCheckbox.value = true;

    // OK / Cancel
    var btnGroup = dialog.add("group");
    btnGroup.alignment = "center";
    btnGroup.spacing = 10;
    var okBtn = btnGroup.add("button", undefined, "Convert", {name: "ok"});
    var cancelBtn = btnGroup.add("button", undefined, "Cancel", {name: "cancel"});

    if (dialog.show() !== 1) return;

    // Collect selections before we start (dialog is closed)
    var queue = [];
    for (var i = 0; i < checkboxes.length; i++) {
        if (checkboxes[i].value) {
            queue.push({
                doc: checkboxes[i]._doc,
                colorMode: checkboxes[i]._colorMode,
                hasPath: checkboxes[i]._hasPath
            });
        }
    }

    if (queue.length === 0) {
        alert("No documents selected.");
        return;
    }

    var doSave = saveCheckbox.value;

    // ---------- Conversion ----------
    // Timing notes:
    //   - executeMenuCommand("doc-color-cmyk") is synchronous, but Illustrator
    //     needs a beat afterward to finish re-processing swatches/gradients.
    //   - We $.sleep() between steps and call app.redraw() so subsequent
    //     save() calls don't race the color-mode change.

    var converted = 0;
    var alreadyCMYK = 0;
    var savedCount = 0;
    var skippedUnsaved = 0;
    var errors = [];

    for (var i = 0; i < queue.length; i++) {
        var entry = queue[i];
        var doc = entry.doc;
        var docName = "Document " + (i + 1);
        try { docName = doc.name; } catch (e) {}

        try {
            // Activate this document
            app.activeDocument = doc;
            $.sleep(150);
            app.redraw();

            // Re-read the current color space from the now-active document
            var activeDoc = app.activeDocument;

            if (activeDoc.documentColorSpace === DocumentColorSpace.CMYK) {
                alreadyCMYK++;
                if (doSave && entry.hasPath) {
                    activeDoc.save();
                    $.sleep(200);
                    savedCount++;
                }
                continue;
            }

            // Perform the conversion
            app.executeMenuCommand("doc-color-cmyk");
            $.sleep(400);       // let Illustrator finish the conversion
            app.redraw();
            $.sleep(150);

            // Re-grab the active document reference after the mode switch
            activeDoc = app.activeDocument;

            // Verify
            if (activeDoc.documentColorSpace !== DocumentColorSpace.CMYK) {
                errors.push(docName + ": conversion did not complete");
                continue;
            }
            converted++;

            // Save
            if (doSave) {
                if (entry.hasPath) {
                    activeDoc.save();
                    $.sleep(250);
                    savedCount++;
                } else {
                    skippedUnsaved++;
                }
            }
        } catch (e) {
            errors.push(docName + ": " + e.message);
        }
    }

    // ---------- Summary ----------
    var lines = [];
    lines.push("Batch Conversion Complete");
    lines.push("");
    lines.push("Converted to CMYK: " + converted);
    if (alreadyCMYK > 0)      lines.push("Already CMYK: " + alreadyCMYK);
    if (doSave)               lines.push("Saved: " + savedCount);
    if (skippedUnsaved > 0)   lines.push("Skipped save (no file path): " + skippedUnsaved);
    if (errors.length > 0) {
        lines.push("");
        lines.push("Errors:");
        for (var i = 0; i < errors.length; i++) lines.push("  \u2022 " + errors[i]);
    }

    // Use a dialog for the report so long lists don't get clipped
    var report = new Window("dialog", "RGB to CMYK Batch");
    report.orientation = "column";
    report.alignChildren = "fill";
    report.margins = 16;
    report.spacing = 10;

    var txt = report.add("edittext", undefined, lines.join("\n"),
        {multiline: true, scrolling: true, readonly: true});
    txt.preferredSize.width = 440;
    txt.preferredSize.height = 220;

    var closeGroup = report.add("group");
    closeGroup.alignment = "center";
    var closeBtn = closeGroup.add("button", undefined, "Close", {name: "ok"});
    closeBtn.onClick = function() { report.close(); };

    report.show();
})();
