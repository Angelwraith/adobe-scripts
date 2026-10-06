/*@METADATA{
  "name": "Open PDF with Comments",
  "description": "Pick a PDF, place every page as a linked file on its own artboard, and redraw its Acrobat comments as editable art on a top 'Acrobat Notes' layer.",
  "version": "3.1",
  "target": "illustrator",
  "tags": ["pdf", "acrobat", "comments", "import", "artboard", "notes"]
}@END_METADATA*/

/*
    Open PDF with Comments.js
    ---------------------------------------------------------------
    1. Pick a PDF.
    2. Acrobat (running hidden) reads every comment - type, position,
       colors, text - and hands it back to Illustrator.
    3. A new document is created with one artboard per page. Each page is
       PLACED AS A LINK (so Embed / Flatten Transparency still work).
    4. Comments are redrawn as native Illustrator objects on a separate
       top-level "Acrobat Notes" layer, grouped per page - text boxes,
       sticky notes, rectangles, ovals, lines, arrows, polygons, pencil
       marks, highlights/underlines/strikeouts.
    5. Stamps (and anything else that cannot be redrawn) are treated as
       ARTWORK, since applied stamps often carry real content. Acrobat
       deletes the redrawable comments from a copy, flattens what is left,
       and saves "<name> - WITH STAMPS.pdf" next to the original. The pages
       then link to that copy instead. It is only made when such items
       exist; otherwise the pages link to the untouched original.

    Requires: Windows + Acrobat Pro or Standard (not Reader).
*/

#target illustrator

(function () {

    // ---------------- settings ----------------
    var SUFFIX          = " - WITH STAMPS";
    var COLUMNS         = 4;          // artboards per row
    var GAP             = 72;         // points between artboards (72 = 1 in)
    var TIMEOUT_SEC     = 180;        // how long to wait for Acrobat
    var PAGE_BOX        = PDFBoxType.PDFCROPBOX; // keep as Crop box - note positions are measured from it
    var PAGES_LAYER     = "PDF Pages";
    var NOTES_LAYER     = "Acrobat Notes";
    var NOTES_PRINTABLE = false;      // false = notes layer will not print
    var NOTE_FONTS      = ["Helvetica", "ArialMT", "MyriadPro-Regular"]; // first one found is used
    var DEFAULT_COLOR   = [255, 0, 0]; // used when a comment has no color
    // ------------------------------------------

    if ($.os.toLowerCase().indexOf("windows") === -1) {
        alert("This script is Windows-only (it drives Acrobat through Windows scripting).");
        return;
    }

    var src = File.openDialog("Select a PDF with Acrobat comments", "PDF files:*.pdf");
    if (!src) return;

    var baseName = decodeURI(src.name).replace(/\.pdf$/i, "");
    var flatFile = new File(src.parent.fsName + "\\" + baseName + SUFFIX + ".pdf");
    var linkFile = src;

    // ---------- 1. Read comments via Acrobat ----------
    var info = runAcrobat(src, flatFile, TIMEOUT_SEC);
    if (!info) return; // error already shown
    if (info.pages < 1) { alert("Acrobat reported 0 pages."); return; }
    if (info.stampsBaked && flatFile.exists) linkFile = flatFile;
    else info.stampsBaked = false;

    // ---------- 2. Build Illustrator document ----------
    var oldUIL    = app.userInteractionLevel;
    var oldCoords = app.coordinateSystem;
    var pdfOpts   = app.preferences.PDFFileOptions;
    var oldPage   = pdfOpts.pageToOpen;
    var oldBox    = pdfOpts.pDFCropToBox;

    var drawn = 0, baked = 0, failed = 0, rotatedSkipped = 0;

    try {
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;
        app.coordinateSystem     = CoordinateSystem.DOCUMENTCOORDINATESYSTEM;
        pdfOpts.pDFCropToBox     = PAGE_BOX;

        var doc        = app.documents.add(DocumentColorSpace.CMYK, 612, 792);
        var pagesLayer = doc.layers[0];
        pagesLayer.name = PAGES_LAYER;

        var r0      = doc.artboards[0].artboardRect; // [L, T, R, B]
        var originX = r0[0];
        var originY = r0[1];

        var pagePos = {}; // page index (0-based) -> [L, T]
        var x = 0, y = 0, rowH = 0, col = 0;

        for (var p = 1; p <= info.pages; p++) {
            pdfOpts.pageToOpen = p;

            var item  = pagesLayer.placedItems.add();
            item.file = linkFile;            // original PDF, or the stamps-baked copy
            item.name = "Page " + p;

            var w = item.width, h = item.height;

            if (col === COLUMNS) { col = 0; x = 0; y += rowH + GAP; rowH = 0; }

            var L = originX + x;
            var T = originY - y;
            item.position = [L, T];
            pagePos[p - 1] = [L, T];

            var gb = item.geometricBounds;
            var ab;
            if (p === 1) { ab = doc.artboards[0]; ab.artboardRect = [gb[0], gb[1], gb[2], gb[3]]; }
            else         { ab = doc.artboards.add([gb[0], gb[1], gb[2], gb[3]]); }
            ab.name = "Page " + p;

            x    += w + GAP;
            rowH  = Math.max(rowH, h);
            col++;
        }

        // ---------- 3. Notes layer ----------
        if (info.annots.length > 0) {
            var notesLayer = doc.layers.add();   // new layers go on top
            notesLayer.name = NOTES_LAYER;
            var font = findFont(NOTE_FONTS);
            var pageGroups = {};

            for (var i = 0; i < info.annots.length; i++) {
                var an = info.annots[i];
                var pg = info.pageInfo[an.page];
                if (!pg || !pagePos[an.page]) { failed++; continue; }
                if (pg.rot !== 0) { rotatedSkipped++; continue; }

                if (!pageGroups[an.page]) {
                    pageGroups[an.page] = notesLayer.groupItems.add();
                    pageGroups[an.page].name = "Page " + (an.page + 1) + " Notes";
                }
                var grp = pageGroups[an.page];
                var map = makeMapper(pagePos[an.page], pg.crop);

                try {
                    var how = drawAnnot(grp, an, map, font);
                    if (how === "baked") baked++; else drawn++;
                } catch (e) {
                    failed++;
                    $.writeln("[WARN] Could not draw " + an.type + " on page " + (an.page + 1) + ": " + e);
                }
            }
            for (var gk in pageGroups) {            // drop groups left empty by baked stamps
                if (pageGroups[gk].pageItems.length === 0) pageGroups[gk].remove();
            }
            notesLayer.printable = NOTES_PRINTABLE;
            if (notesLayer.pageItems.length === 0) notesLayer.remove();
        }

        doc.selection = null;
        doc.artboards.setActiveArtboardIndex(0);
        try { app.executeMenuCommand("fitall"); } catch (e2) {}

    } catch (err) {
        alert("Error while building the document:\n" + err + (err.line ? "\nLine " + err.line : ""));
    } finally {
        pdfOpts.pageToOpen       = oldPage;
        pdfOpts.pDFCropToBox     = oldBox;
        app.coordinateSystem     = oldCoords;
        app.userInteractionLevel = oldUIL;
    }

    // ---------- 4. Summary ----------
    var msg = info.pages + " page(s) placed as links to:\n" + decodeURI(linkFile.name) + "\n\n";
    if (info.annots.length === 0) {
        msg += "No comments found in this PDF.";
    } else {
        msg += drawn + " comment(s) redrawn on the \"" + NOTES_LAYER + "\" layer.";
        if (baked)          msg += "\n" + baked + " stamp(s)/other item(s) baked into the linked copy.";
        if (info.stampCount && !info.stampsBaked) {
            msg += "\n[WARN] " + info.stampCount + " stamp(s) could not be baked in - shown as dashed boxes.";
        }
        if (rotatedSkipped) msg += "\n[WARN] " + rotatedSkipped + " comment(s) skipped on rotated pages.";
        if (failed)         msg += "\n[WARN] " + failed + " comment(s) could not be drawn (see ExtendScript console).";
    }
    alert(msg);


    // =====================================================================
    // DRAWING
    // =====================================================================

    // Converts PDF page coordinates to document coordinates.
    function makeMapper(pos, crop) {
        // crop = [left, top, right, bottom] in PDF space (y up)
        return {
            x: function (px) { return pos[0] + (px - crop[0]); },
            y: function (py) { return pos[1] - (crop[1] - py); }
        };
    }

    // Returns "draw" or "clip".
    function drawAnnot(grp, an, map, font) {
        var r = normRect(an.rect);                  // [x1, y1, x2, y2] PDF, y up
        var left = map.x(r[0]), right = map.x(r[2]);
        var top  = map.y(r[3]), bottom = map.y(r[1]);
        var w = right - left, h = top - bottom;
        var sw = an.width > 0 ? an.width : 0;
        var stroke = makeColor(an.stroke);
        var fill   = makeColor(an.fill);
        var label  = an.type + (an.contents ? ": " + an.contents.replace(/\\n/g, " ").substr(0, 40) : "");
        var item;

        switch (an.type) {

            case "Square":
            case "Circle":
                var inset = sw / 2;
                item = (an.type === "Square")
                    ? grp.pathItems.rectangle(top - inset, left + inset, w - sw, h - sw)
                    : grp.pathItems.ellipse(top - inset, left + inset, w - sw, h - sw);
                styleShape(item, stroke, fill, sw || 1);
                break;

            case "Line":
            case "PolyLine":
            case "Polygon":
                var pts = mapPoints(an.geom, map);
                if (pts.length < 2) throw new Error("no points");
                item = grp.pathItems.add();
                item.setEntirePath(pts);
                item.closed = (an.type === "Polygon");
                styleShape(item, stroke, (an.type === "Polygon") ? fill : null, sw || 1);
                break;

            case "Ink":
                item = grp.groupItems.add();
                var strokes = an.geom.split(";");
                for (var s = 0; s < strokes.length; s++) {
                    var ip = mapPoints(strokes[s], map);
                    if (ip.length < 2) continue;
                    var pth = item.pathItems.add();
                    pth.setEntirePath(ip);
                    pth.closed = false;
                    styleShape(pth, stroke, null, sw || 1);
                }
                break;

            case "Highlight":
                item = grp.pathItems.rectangle(top, left, w, h);
                styleShape(item, null, stroke || colorFromArray([255, 255, 0]), 0);
                item.opacity = 40;
                break;

            case "Underline":
            case "Squiggly":
            case "StrikeOut":
                var ly = (an.type === "StrikeOut") ? (top + bottom) / 2 : bottom + 1;
                item = grp.pathItems.add();
                item.setEntirePath([[left, ly], [right, ly]]);
                styleShape(item, stroke, null, 1);
                break;

            case "FreeText":
                item = grp.groupItems.add();
                if (sw > 0 && stroke) {
                    var box = item.pathItems.rectangle(top - sw / 2, left + sw / 2, w - sw, h - sw);
                    styleShape(box, stroke, fill, sw);
                } else if (fill) {
                    var bg = item.pathItems.rectangle(top, left, w, h);
                    styleShape(bg, null, fill, 0);
                }
                var tf = item.textFrames.add();
                tf.contents = an.contents.replace(/\\n/g, "\r");
                var ca = tf.textRange.characterAttributes;
                ca.size = an.textSize > 0 ? an.textSize : 12;
                ca.fillColor = makeColor(an.textColor) || stroke || colorFromArray(DEFAULT_COLOR);
                if (font) ca.textFont = font;
                tf.position = [left + 2 + sw, top - 2 - sw];
                break;

            case "Text": // sticky note: colored square + its text beside it
                item = grp.groupItems.add();
                var icon = item.pathItems.rectangle(top, left, 14, 14);
                styleShape(icon, colorFromArray([0, 0, 0]), stroke || colorFromArray([255, 230, 0]), 0.5);
                var body = (an.author ? an.author + ": " : "") + an.contents.replace(/\\n/g, "\r");
                var nt = item.textFrames.add();
                nt.contents = body || "(empty note)";
                nt.textRange.characterAttributes.size = 10;
                nt.textRange.characterAttributes.fillColor = colorFromArray(DEFAULT_COLOR);
                if (font) nt.textRange.characterAttributes.textFont = font;
                nt.position = [left + 18, top];
                break;

            default: // Stamp, FileAttachment, Caret, etc.
                // These are baked into the linked page copy by Acrobat, so
                // nothing is drawn here unless that copy could not be made.
                if (info.stampsBaked) return "baked";
                item = grp.groupItems.add();
                var db = item.pathItems.rectangle(top, left, w, h);
                styleShape(db, stroke || colorFromArray(DEFAULT_COLOR), null, 1);
                db.strokeDashes = [4, 3];
                var lt = item.textFrames.add();
                lt.contents = "[" + an.type + "]";
                lt.textRange.characterAttributes.size = 9;
                lt.position = [left, top + 12];
                break;
        }

        if (an.opacity > 0 && an.opacity < 1 && an.type !== "Highlight") item.opacity = an.opacity * 100;
        item.name = label;
        return "draw";
    }

    function styleShape(item, stroke, fill, sw) {
        if (stroke) { item.stroked = true; item.strokeColor = stroke; item.strokeWidth = sw; }
        else        { item.stroked = false; }
        if (fill)   { item.filled = true; item.fillColor = fill; }
        else        { item.filled = false; }
    }

    function normRect(r) {
        return [Math.min(r[0], r[2]), Math.min(r[1], r[3]), Math.max(r[0], r[2]), Math.max(r[1], r[3])];
    }

    function mapPoints(str, map) {
        var n = str ? str.split(",") : [], out = [];
        for (var k = 0; k + 1 < n.length; k += 2) {
            out.push([map.x(parseFloat(n[k])), map.y(parseFloat(n[k + 1]))]);
        }
        return out;
    }

    // Acrobat color string -> Illustrator color (or null for transparent/none)
    function makeColor(str) {
        if (!str) return null;
        var c = str.split(",");
        var sp = c[0];
        if (sp === "RGB" && c.length >= 4) {
            return colorFromArray([c[1] * 255, c[2] * 255, c[3] * 255]);
        }
        if (sp === "G" && c.length >= 2) {
            var g = new GrayColor(); g.gray = (1 - c[1]) * 100; return g;
        }
        if (sp === "CMYK" && c.length >= 5) {
            var k = new CMYKColor();
            k.cyan = c[1] * 100; k.magenta = c[2] * 100; k.yellow = c[3] * 100; k.black = c[4] * 100;
            return k;
        }
        return null; // "T" = transparent
    }

    function colorFromArray(a) {
        var c = new RGBColor(); c.red = a[0]; c.green = a[1]; c.blue = a[2]; return c;
    }

    function findFont(names) {
        for (var f = 0; f < names.length; f++) {
            try { return app.textFonts.getByName(names[f]); } catch (e) {}
        }
        return null;
    }


    // =====================================================================
    // ACROBAT
    // Writes a small VBScript that drives Acrobat via COM, runs it, and
    // waits for its result file. Returns
    //   { pages, stampsBaked, stampCount, pageInfo{idx:{crop,rot}}, annots[] } or null.
    // =====================================================================
    function runAcrobat(srcFile, dstFile, timeoutSec) {
        var stamp   = new Date().getTime();
        var vbsFile = new File(Folder.temp.fsName + "\\ai_notes_" + stamp + ".vbs");
        var outFile = new File(Folder.temp.fsName + "\\ai_notes_" + stamp + ".txt");
        var tmpFile = new File(outFile.fsName + ".part");

        function q(s) { return '"' + String(s).replace(/"/g, '""') + '"'; }

        var vbs = [
            'On Error Resume Next',
            'Dim fso, app, pd, js, n, i, j, a, an, buf, needFlat, supported, fl, stampCount, tries',
            'Dim t, pg, rc, sc, fc, tc, wd, ts, op, au, ct, ge, g',
            'Set fso = CreateObject("Scripting.FileSystemObject")',
            'supported = "|FreeText|Text|Square|Circle|Line|Polygon|PolyLine|Ink|Highlight|Underline|Squiggly|StrikeOut|"',
            '',
            'Sub Finish(msg)',
            '  Dim out',
            '  Set out = fso.CreateTextFile(' + q(tmpFile.fsName) + ', True, True)',
            '  out.Write msg',
            '  out.Close',
            '  fso.MoveFile ' + q(tmpFile.fsName) + ', ' + q(outFile.fsName),
            '  WScript.Quit',
            'End Sub',
            '',
            'Function Flat(v)',
            '  Dim e, s',
            '  If IsArray(v) Then',
            '    s = ""',
            '    For Each e In v',
            '      If s <> "" Then s = s & ","',
            '      s = s & Flat(e)',
            '    Next',
            '    Flat = s',
            '  ElseIf IsNull(v) Or IsEmpty(v) Then',
            '    Flat = ""',
            '  ElseIf VarType(v) = vbString Then',
            '    Flat = v',
            '  Else',
            '    Flat = Replace(CStr(v), ",", ".")',
            '  End If',
            'End Function',
            '',
            'Function Clean(v)',
            '  Dim s',
            '  s = ""',
            '  If Not (IsNull(v) Or IsEmpty(v)) Then s = CStr(v)',
            '  s = Replace(s, vbCrLf, "\\n")',
            '  s = Replace(s, vbCr, "\\n")',
            '  s = Replace(s, vbLf, "\\n")',
            '  Clean = Replace(s, vbTab, " ")',
            'End Function',
            '',
            '\' Acrobat can be busy (starting up, updating, or left running by an',
            '\' earlier run), so try to attach to a running copy, then retry.',
            'Err.Clear',
            'Set app = GetObject(, "AcroExch.App")',
            'If Err.Number <> 0 Then',
            '  For tries = 1 To 3',
            '    Err.Clear',
            '    Set app = CreateObject("AcroExch.App")',
            '    If Err.Number = 0 Then Exit For',
            '    WScript.Sleep 3000',
            '  Next',
            'End If',
            'If Err.Number <> 0 Then Finish "ERR" & vbTab & "Could not start Acrobat - error " & Hex(Err.Number) & ": " & Err.Description & vbCrLf & _',
            '  "Acrobat Pro or Standard must be installed (Reader will not work). If Acrobat is already open, close it - including any leftover Acrobat processes in Task Manager - and run the script again."',
            'Err.Clear',
            'Set pd = CreateObject("AcroExch.PDDoc")',
            'If Err.Number <> 0 Then Finish "ERR" & vbTab & "Acrobat started but would not create a document object - error " & Hex(Err.Number) & ": " & Err.Description',
            'If Not pd.Open(' + q(srcFile.fsName) + ') Then Finish "ERR" & vbTab & "Acrobat could not open the PDF (is it open in Acrobat already, or on a drive Acrobat cannot reach?)."',
            'n = pd.GetNumPages()',
            'Set js = pd.GetJSObject()',
            'If Err.Number <> 0 Or Not IsObject(js) Then',
            '  pd.Close',
            '  Finish "ERR" & vbTab & "Acrobat would not hand over its scripting object - error " & Hex(Err.Number) & ": " & Err.Description',
            'End If',
            'Err.Clear',
            'buf = ""',
            'For i = 0 To n - 1',
            '  rc = "" : rc = Flat(js.getPageBox("Crop", i))',
            '  g = 0 : g = js.getPageRotation(i)',
            '  buf = buf & "P" & vbTab & i & vbTab & rc & vbTab & g & vbCrLf',
            'Next',
            '',
            'Err.Clear',
            'js.syncAnnotScan',
            'a = Empty',
            'a = js.getAnnots()',
            'needFlat = False',
            'stampCount = 0',
            'If IsArray(a) Then',
            '  For Each an In a',
            '    t = "" : t = an.type',
            '    If t <> "Popup" And t <> "" Then',
            '      pg = -1 : pg = an.page',
            '      rc = "" : rc = Flat(an.rect)',
            '      sc = "" : sc = Flat(an.strokeColor)',
            '      fc = "" : fc = Flat(an.fillColor)',
            '      tc = "" : tc = Flat(an.textColor)',
            '      wd = 0  : wd = an.width',
            '      ts = 0  : ts = an.textSize',
            '      op = 1  : op = an.opacity',
            '      au = "" : au = an.author',
            '      ct = "" : ct = an.contents',
            '      ge = ""',
            '      If t = "Line" Then ge = Flat(an.points)',
            '      If t = "Polygon" Or t = "PolyLine" Then ge = Flat(an.vertices)',
            '      If t = "Ink" Then',
            '        g = Empty : g = an.gestures',
            '        If IsArray(g) Then',
            '          For Each j In g',
            '            If ge <> "" Then ge = ge & ";"',
            '            ge = ge & Flat(j)',
            '          Next',
            '        End If',
            '      End If',
            '      If InStr(supported, "|" & t & "|") = 0 Then',
            '        needFlat = True',
            '        stampCount = stampCount + 1',
            '      End If',
            '      buf = buf & "A" & vbTab & pg & vbTab & t & vbTab & rc & vbTab & sc & vbTab & fc & vbTab & tc & vbTab & _',
            '            Flat(wd) & vbTab & Flat(ts) & vbTab & Flat(op) & vbTab & Clean(au) & vbTab & Clean(ct) & vbTab & ge & vbCrLf',
            '    End If',
            '  Next',
            'End If',
            'Err.Clear',
            '',
            '\' Stamps and anything else we cannot redraw get baked into a copy:',
            '\' delete the redrawable comments, flatten what is left, save as a new file.',
            'If needFlat Then',
            '  If IsArray(a) Then',
            '    For i = UBound(a) To 0 Step -1',
            '      t = "" : t = a(i).type',
            '      If InStr(supported, "|" & t & "|") > 0 Or t = "Popup" Then a(i).destroy',
            '      Err.Clear',
            '    Next',
            '  End If',
            '  js.flattenPages 0, n - 1, 2',
            '  If Err.Number = 0 Then',
            '    If Not pd.Save(1, ' + q(dstFile.fsName) + ') Then needFlat = False',
            '  Else',
            '    needFlat = False',
            '  End If',
            '  Err.Clear',
            'End If',
            'pd.Close',
            'If app.GetNumAVDocs() = 0 Then app.Exit',
            'If needFlat Then fl = "1" Else fl = "0"',
            'Finish "OK" & vbTab & n & vbTab & fl & vbTab & stampCount & vbCrLf & buf'
        ].join("\r\n");

        vbsFile.lineFeed = "Windows";
        if (!vbsFile.open("w")) { alert("Could not write temp script:\n" + vbsFile.fsName); return null; }
        vbsFile.write(vbs);
        vbsFile.close();

        if (outFile.exists) outFile.remove();
        if (tmpFile.exists) tmpFile.remove();
        vbsFile.execute();

        var waited = 0;
        while (!outFile.exists && waited < timeoutSec * 1000) {
            $.sleep(250);
            waited += 250;
        }

        if (!outFile.exists) {
            alert("Timed out waiting for Acrobat (" + timeoutSec + "s).\n" +
                  "Check for an Acrobat dialog that needs attention.");
            return null;
        }

        outFile.encoding = "UTF-16";   // VBScript writes UTF-16LE with a BOM
        outFile.open("r");
        var txt = outFile.read();
        outFile.close();
        try { outFile.remove(); vbsFile.remove(); } catch (e) {}

        txt = txt.replace(/^\uFEFF/, "");
        var lines = txt.split(/\r\n|\n/);
        var head  = lines[0].split("\t");
        if (head[0] !== "OK") {
            alert("Acrobat step failed:\n\n" + txt.replace(/^ERR\t/, ""));
            return null;
        }

        var res = {
            pages: parseInt(head[1], 10),
            stampsBaked: head[2] === "1",
            stampCount: parseInt(head[3], 10) || 0,
            pageInfo: {},
            annots: []
        };

        for (var li = 1; li < lines.length; li++) {
            var f = lines[li].split("\t");
            if (f[0] === "P") {
                var cb = f[2].split(",");
                // Acrobat crop box = [left, top, right, bottom]
                res.pageInfo[parseInt(f[1], 10)] = {
                    crop: [parseFloat(cb[0]), parseFloat(cb[1]), parseFloat(cb[2]), parseFloat(cb[3])],
                    rot: parseInt(f[3], 10) || 0
                };
            } else if (f[0] === "A" && f.length >= 13) {
                var rr = f[3].split(",");
                if (rr.length < 4) continue;
                res.annots.push({
                    page: parseInt(f[1], 10),
                    type: f[2],
                    rect: [parseFloat(rr[0]), parseFloat(rr[1]), parseFloat(rr[2]), parseFloat(rr[3])],
                    stroke: f[4], fill: f[5], textColor: f[6],
                    width: parseFloat(f[7]) || 0,
                    textSize: parseFloat(f[8]) || 0,
                    opacity: parseFloat(f[9]) || 1,
                    author: f[10], contents: f[11], geom: f[12]
                });
            }
        }
        return res;
    }

})();
