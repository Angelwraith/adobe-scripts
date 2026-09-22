/*@METADATA{
  "name": "Open PDF with Comments",
  "description": "Pick a PDF, have Acrobat flatten its comments into a copy, then place every page as a linked file on its own artboard.",
  "version": "1.0",
  "target": "illustrator",
  "tags": ["pdf", "acrobat", "comments", "import", "artboard"]
}@END_METADATA*/

/*
    Open PDF with Comments.js
    ---------------------------------------------------------------
    1. Pick a PDF.
    2. Acrobat (running hidden) flattens all comments/markups into page
       content and saves "<name> - NOTES FLATTENED.pdf" next to the original.
       (If the PDF has no comments, the original is used as-is.)
    3. A new Illustrator document is created with one artboard per page,
       each page PLACED AS A LINK - so you can Embed or Flatten
       Transparency afterward, whichever behaves.

    Requires: Windows + Acrobat Pro or Standard (not Reader).
*/

#target illustrator

(function () {

    // ---------------- settings ----------------
    var SUFFIX       = " - NOTES FLATTENED";
    var COLUMNS      = 4;        // artboards per row
    var GAP          = 72;       // points between artboards (72 = 1 in)
    var TIMEOUT_SEC  = 180;      // how long to wait for Acrobat
    var PAGE_BOX     = PDFBoxType.PDFCROPBOX; // PDFTRIMBOX / PDFMEDIABOX / PDFBLEEDBOX / PDFARTBOX
    var LAYER_NAME   = "PDF Pages";
    // ------------------------------------------

    if ($.os.toLowerCase().indexOf("windows") === -1) {
        alert("This script is Windows-only (it drives Acrobat through Windows scripting).");
        return;
    }

    var src = File.openDialog("Select a PDF with Acrobat comments", "PDF files:*.pdf");
    if (!src) return;

    // ---------- 1. Flatten comments via Acrobat ----------
    var baseName = decodeURI(src.name).replace(/\.pdf$/i, "");
    var dst      = new File(src.parent.fsName + "\\" + baseName + SUFFIX + ".pdf");

    var result = runAcrobatFlatten(src, dst, TIMEOUT_SEC);
    if (!result) return; // error already shown

    var pdfToPlace = result.annots > 0 ? dst : src;
    var pageCount  = result.pages;
    if (pageCount < 1) { alert("Acrobat reported 0 pages."); return; }

    // ---------- 2. Build Illustrator document ----------
    var oldUIL    = app.userInteractionLevel;
    var oldCoords = app.coordinateSystem;
    var pdfOpts   = app.preferences.PDFFileOptions;
    var oldPage   = pdfOpts.pageToOpen;
    var oldBox    = pdfOpts.pDFCropToBox;

    try {
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;
        app.coordinateSystem     = CoordinateSystem.DOCUMENTCOORDINATESYSTEM;
        pdfOpts.pDFCropToBox     = PAGE_BOX;

        var doc   = app.documents.add(DocumentColorSpace.CMYK, 612, 792);
        var layer = doc.layers[0];
        layer.name = LAYER_NAME;

        var r0      = doc.artboards[0].artboardRect; // [L, T, R, B]
        var originX = r0[0];
        var originY = r0[1];

        var x = 0, y = 0, rowH = 0, col = 0;

        for (var p = 1; p <= pageCount; p++) {
            pdfOpts.pageToOpen = p;

            var item  = layer.placedItems.add();
            item.file = pdfToPlace;          // linked, not embedded
            item.name = "Page " + p;

            var w = item.width, h = item.height;

            if (col === COLUMNS) {           // new row
                col = 0; x = 0; y += rowH + GAP; rowH = 0;
            }

            var L = originX + x;
            var T = originY - y;
            item.position = [L, T];

            var gb   = item.geometricBounds; // [L, T, R, B]
            var rect = [gb[0], gb[1], gb[2], gb[3]];

            var ab;
            if (p === 1) { ab = doc.artboards[0]; ab.artboardRect = rect; }
            else         { ab = doc.artboards.add(rect); }
            ab.name = "Page " + p;

            x    += w + GAP;
            rowH  = Math.max(rowH, h);
            col++;
        }

        doc.selection = null;
        doc.artboards.setActiveArtboardIndex(0);
        try { app.executeMenuCommand("fitall"); } catch (e) {}

    } catch (err) {
        alert("Error while placing pages:\n" + err + (err.line ? "\nLine " + err.line : ""));
    } finally {
        pdfOpts.pageToOpen       = oldPage;
        pdfOpts.pDFCropToBox     = oldBox;
        app.coordinateSystem     = oldCoords;
        app.userInteractionLevel = oldUIL;
    }

    if (result.annots === 0) {
        alert("No comments found - placed the original PDF (" + pageCount + " pages).");
    }


    // =====================================================================
    // Writes a small VBScript that drives Acrobat via COM, runs it, and waits
    // for a result file.  Returns {pages, annots} or null on failure.
    // =====================================================================
    function runAcrobatFlatten(srcFile, dstFile, timeoutSec) {
        var stamp   = new Date().getTime();
        var vbsFile = new File(Folder.temp.fsName + "\\ai_flatten_" + stamp + ".vbs");
        var outFile = new File(Folder.temp.fsName + "\\ai_flatten_" + stamp + ".txt");

        function q(s) { return '"' + String(s).replace(/"/g, '""') + '"'; }

        var vbs = [
            'On Error Resume Next',
            'Dim fso, out, app, pd, js, n, cnt, a',
            'Set fso = CreateObject("Scripting.FileSystemObject")',
            'Sub Finish(msg)',
            '  Set out = fso.CreateTextFile(' + q(outFile.fsName) + ', True)',
            '  out.Write msg',
            '  out.Close',
            '  WScript.Quit',
            'End Sub',
            '',
            'Set app = CreateObject("AcroExch.App")',
            'If Err.Number <> 0 Then Finish "ERR|Could not start Acrobat (Pro/Standard required - Reader will not work)."',
            'Set pd = CreateObject("AcroExch.PDDoc")',
            'If Not pd.Open(' + q(srcFile.fsName) + ') Then Finish "ERR|Acrobat could not open the PDF."',
            'n = pd.GetNumPages()',
            'Set js = pd.GetJSObject()',
            'Err.Clear',
            'js.syncAnnotScan',
            'cnt = 0',
            'a = js.getAnnots()',
            'If Err.Number = 0 Then',
            '  If IsArray(a) Then cnt = UBound(a) + 1',
            'End If',
            'Err.Clear',
            'If cnt > 0 Then',
            '  js.flattenPages 0, n - 1, 2',   // 2 = flatten non-printing markups too
            '  If Err.Number <> 0 Then',
            '    pd.Close',
            '    Finish "ERR|Acrobat failed to flatten comments: " & Err.Description',
            '  End If',
            '  If Not pd.Save(1, ' + q(dstFile.fsName) + ') Then',
            '    pd.Close',
            '    Finish "ERR|Could not save flattened PDF (is it open somewhere?)."',
            '  End If',
            'End If',
            'pd.Close',
            'If app.GetNumAVDocs() = 0 Then app.Exit',
            'Finish "OK|" & n & "|" & cnt'
        ].join("\r\n");

        vbsFile.lineFeed = "Windows";
        if (!vbsFile.open("w")) { alert("Could not write temp script:\n" + vbsFile.fsName); return null; }
        vbsFile.write(vbs);
        vbsFile.close();

        if (outFile.exists) outFile.remove();
        vbsFile.execute();

        // wait for result
        var waited = 0;
        while (!outFile.exists && waited < timeoutSec * 1000) {
            $.sleep(250);
            waited += 250;
        }
        $.sleep(250); // let the file finish writing

        if (!outFile.exists) {
            alert("Timed out waiting for Acrobat (" + timeoutSec + "s).\n" +
                  "Check for an Acrobat dialog that needs attention.");
            return null;
        }

        outFile.open("r");
        var txt = outFile.read();
        outFile.close();
        try { outFile.remove(); vbsFile.remove(); } catch (e) {}

        var parts = txt.split("|");
        if (parts[0] !== "OK") {
            alert("Acrobat step failed:\n" + (parts[1] || txt));
            return null;
        }
        return { pages: parseInt(parts[1], 10), annots: parseInt(parts[2], 10) };
    }

})();
