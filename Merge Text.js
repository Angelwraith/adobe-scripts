/*@METADATA{
  "name": "Merge Text",
  "description": "Merges selected text frames into one, preserving character formatting and exact positions. Frames side-by-side join on the same line (spacing kept via tracking); stacked frames join as new lines (spacing kept via leading).",
  "version": "1.2",
  "target": "illustrator",
  "tags": ["text", "merge", "combine", "labels"]
}@END_METADATA*/

#target illustrator

// ============================================================================
// Merge Text
// ----------------------------------------------------------------------------
//   1. Groups selected frames into ROWS (frames that overlap vertically).
//   2. Each row is joined left-to-right on ONE line. The horizontal gap between
//      frames is reproduced with tracking on the join character (measured and
//      corrected iteratively), and baseline differences via baseline shift.
//   3. Rows are then joined top-to-bottom with paragraph breaks. The
//      baseline-to-baseline distance is applied as leading on each new line.
//   4. textRange.move() is used so kerning/tracking/color/fonts survive.
// ============================================================================

(function () {
    if (app.documents.length === 0) { alert("Please open a document first."); return; }

    var doc = app.activeDocument;
    var sel = doc.selection;
    var TOL = 0.01;

    var tfs = [];
    for (var i = 0; i < sel.length; i++) {
        if (sel[i] && sel[i].typename === "TextFrame" && sel[i].characters.length > 0) tfs.push(sel[i]);
    }
    if (tfs.length < 2) { alert("Select two or more text frames to merge."); return; }

    function isPoint(f) { return f.kind === TextType.POINTTEXT; }
    function baseline(f) { return isPoint(f) ? f.anchor[1] : f.top; }

    // Snapshot geometry BEFORE anything moves
    var items = [];
    for (var i = 0; i < tfs.length; i++) {
        var b = tfs[i].geometricBounds; // [left, top, right, bottom]
        items.push({ f: tfs[i], l: b[0], t: b[1], r: b[2], btm: b[3], base: baseline(tfs[i]) });
    }

    // ---- Group into rows (>= 50% vertical overlap with the row's first frame)
    items.sort(function (a, b) { return b.t - a.t; });
    var rows = [];
    for (var i = 0; i < items.length; i++) {
        var it = items[i], placed = false;
        for (var r = 0; r < rows.length && !placed; r++) {
            var ref = rows[r][0];
            var overlap = Math.min(it.t, ref.t) - Math.max(it.btm, ref.btm);
            var h = Math.min(it.t - it.btm, ref.t - ref.btm);
            if (h > 0 && overlap >= h * 0.5) { rows[r].push(it); placed = true; }
        }
        if (!placed) rows.push([it]);
    }
    for (var r = 0; r < rows.length; r++) rows[r].sort(function (a, b) { return a.l - b.l; });
    rows.sort(function (a, b) { return b[0].base - a[0].base; });

    // ---- Horizontal merge within each row
    var rowFrames = [];
    for (var r = 0; r < rows.length; r++) {
        var row = rows[r];
        var target = row[0].f;
        var rowLeft = row[0].l, rowBase = row[0].base;

        for (var j = 1; j < row.length; j++) {
            var src = row[j];
            var joinIdx = target.characters.length - 1;
            var start = joinIdx + 1;

            src.f.textRange.move(target, ElementPlacement.PLACEATEND);
            var end = target.characters.length;

            // Keep original baseline offset
            var dy = src.base - rowBase;
            if (isPoint(target) && isPoint(src.f) && Math.abs(dy) > TOL) {
                for (var c = start; c < end; c++) {
                    var ca = target.characters[c].characterAttributes;
                    ca.baselineShift = ca.baselineShift + dy;
                }
            }

            // Keep original horizontal gap: tune tracking on the join char
            if (isPoint(target)) {
                var expected = src.r - rowLeft;
                for (var k = 0; k < 6; k++) {
                    var gb = target.geometricBounds;
                    var err = expected - (gb[2] - gb[0]);
                    if (Math.abs(err) < TOL) break;
                    var jca = target.characters[joinIdx].characterAttributes;
                    var hs = (jca.horizontalScale || 100) / 100;
                    jca.tracking = jca.tracking + (err / (jca.size * hs)) * 1000;
                }
            }
            src.f.remove();
        }

        // Put the row back where it started (matters for centered/right text)
        if (isPoint(target)) {
            var nb = target.geometricBounds;
            target.translate(rowLeft - nb[0], rowBase - target.anchor[1]);
        }
        rowFrames.push({ f: target, base: baseline(target) });
    }

    // ---- Vertical merge of rows
    var target = rowFrames[0].f;
    for (var i = 1; i < rowFrames.length; i++) {
        var src = rowFrames[i];
        var gap = rowFrames[i - 1].base - src.base; // baseline-to-baseline

        var oldLen = target.characters.length;
        target.textRange.characters.add("\r");
        src.f.textRange.move(target, ElementPlacement.PLACEATEND);

        var newStart = oldLen + 1;
        var newEnd = target.characters.length;
        var lineEnd = newEnd;
        for (var c = newStart; c < newEnd; c++) {
            if (target.characters[c].contents === "\r") { lineEnd = c; break; }
        }
        // Leading = max over the line's chars, so set it on every char of the
        // new first line (NOT on the \r -- that belongs to the previous line).
        for (var c = newStart; c < lineEnd; c++) {
            var ca = target.characters[c].characterAttributes;
            ca.autoLeading = false;
            ca.leading = gap;
        }
        src.f.remove();
    }

    doc.selection = [target];
})();
