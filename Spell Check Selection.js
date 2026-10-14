/*@METADATA{
  "name": "Spell Check Selection",
  "description": "Runs Illustrator's Check Spelling on the selected text only. Copies selected text into a temp document, spell checks it there, then writes only the corrected words back to the originals (formatting preserved).",
  "version": "1.0",
  "target": "illustrator",
  "tags": ["text", "spelling", "spell check", "selection", "proof"]
}@END_METADATA*/

#target illustrator

// ============================================================================
// Spell Check Selection
// ----------------------------------------------------------------------------
//   PHASE 1 (run with text selected in your working document):
//     - Collects every text story in the selection (groups are searched).
//     - Creates a temp document holding ONLY that text, one frame per story.
//     - Opens Edit > Spelling > Check Spelling on the temp document.
//     - Asks to apply. Yes = corrections are written back, temp doc closes.
//
//   PHASE 2 (only needed if you answered "No" / weren't done):
//     - Run this script again while the temp document is the active doc.
//       It applies the corrections and closes the temp document.
//
//   Write-back is a word-level diff, so only the words that changed are
//   replaced; surrounding character formatting is untouched. Each changed
//   word inherits the formatting of its own first character.
//   Requires Illustrator 2020+ (pageItem.uuid).
// ============================================================================

(function () {
    var TAG_LAYER = "SpellCheckSel|";   // temp doc layer name prefix + source doc name
    var TAG_FRAME = "SC|";              // temp frame name prefix + source item uuid

    if (app.documents.length === 0) { alert("Please open a document first."); return; }
    var doc = app.activeDocument;

    var tagLayer = findTagLayer(doc);
    if (tagLayer) { applyPhase(doc, tagLayer); return; }
    checkPhase(doc);

    // ------------------------------------------------------------------ PHASE 1
    function checkPhase(src) {
        var sel = src.selection;
        var frames = [];

        if (sel && sel.typename === "TextRange") {
            // Cursor / highlighted text inside a frame
            try { frames.push(sel.story.textFrames[0]); } catch (e) {}
        } else if (sel && sel.length) {
            for (var i = 0; i < sel.length; i++) collectFrames(sel[i], frames);
        }
        if (!frames.length) { alert("Select the text (or groups containing text) you want to spell check."); return; }

        // One entry per story (threaded frames share a story)
        var stories = [], seen = {};
        for (var i = 0; i < frames.length; i++) {
            var f = frames[i], st, head, id;
            try { st = f.story; head = st.textFrames[0]; id = head.uuid; } catch (e) {
                alert("This script needs Illustrator 2020 or newer (pageItem.uuid)."); return;
            }
            if (seen[id]) continue;
            seen[id] = true;
            var txt = st.textRange.contents;
            if (!/[A-Za-z]/.test(txt)) continue;
            stories.push({ uuid: id, text: txt });
        }
        if (!stories.length) { alert("No checkable text found in the selection."); return; }

        // Temp document
        var tmp = app.documents.add(DocumentColorSpace.CMYK, 1224, 1584);
        var lay = tmp.layers[0];
        lay.name = TAG_LAYER + src.name;
        var y = -36;
        for (var i = 0; i < stories.length; i++) {
            var t = lay.textFrames.add();
            t.contents = stories[i].text;
            t.name = TAG_FRAME + stories[i].uuid;
            t.position = [36, y];
            y -= (t.height + 24);
        }

        app.redraw();
        try { app.executeMenuCommand("Check Spelling"); } catch (e) {}

        var go = confirm(
            "Apply spelling corrections back to \"" + src.name + "\"?\n\n" +
            "Yes = apply now and close the temp document.\n" +
            "No = keep checking. When done, run this script again\n" +
            "with the temp document active."
        );
        if (go) applyPhase(tmp, lay);
    }

    // ------------------------------------------------------------------ PHASE 2
    function applyPhase(tmp, lay) {
        var srcName = lay.name.substring(TAG_LAYER.length);
        var src = null;
        for (var i = 0; i < app.documents.length; i++) {
            if (app.documents[i].name === srcName) { src = app.documents[i]; break; }
        }
        if (!src) { alert("Source document \"" + srcName + "\" is not open. Open it and run again."); return; }

        // Read everything from the temp doc first
        var jobs = [];
        for (var i = 0; i < lay.textFrames.length; i++) {
            var t = lay.textFrames[i];
            if (t.name.indexOf(TAG_FRAME) !== 0) continue;
            jobs.push({ uuid: t.name.substring(TAG_FRAME.length), text: t.contents });
        }

        app.activeDocument = src;
        var changedStories = 0, changedWords = 0, missing = 0, skipped = 0;

        for (var i = 0; i < jobs.length; i++) {
            var item = null;
            try { item = src.getPageItemFromUuid(jobs[i].uuid); } catch (e) {}
            if (!item) { missing++; continue; }
            var st = item.story;
            var oldText = st.textRange.contents;
            var newText = jobs[i].text;
            if (oldText === newText) continue;
            if (st.characters.length !== oldText.length) { skipped++; continue; }

            var ops = diffOps(oldText, newText);
            for (var k = ops.length - 1; k >= 0; k--) replaceRange(st, ops[k].start, ops[k].len, ops[k].str);
            changedStories++;
            changedWords += ops.length;
        }

        tmp.close(SaveOptions.DONOTSAVECHANGES);
        app.activeDocument = src;

        var msg = changedWords
            ? "Applied " + changedWords + " correction(s) in " + changedStories + " text object(s)."
            : "No corrections to apply.";
        if (missing) msg += "\n" + missing + " original text object(s) could not be found (deleted?).";
        if (skipped) msg += "\n" + skipped + " text object(s) skipped (unexpected character mismatch).";
        alert(msg);
    }

    // ---------------------------------------------------------------- helpers
    function findTagLayer(d) {
        for (var i = 0; i < d.layers.length; i++) {
            if (d.layers[i].name.indexOf(TAG_LAYER) === 0) return d.layers[i];
        }
        return null;
    }

    function collectFrames(it, out) {
        if (!it) return;
        if (it.typename === "TextFrame") { out.push(it); return; }
        if (it.typename === "GroupItem") {
            for (var i = 0; i < it.pageItems.length; i++) collectFrames(it.pageItems[i], out);
        }
    }

    // Replace story chars [start, start+len) with str, keeping formatting of char at start
    function replaceRange(st, start, len, str) {
        if (str.length > 0) {
            st.characters[start].contents = str;   // replaces 1 char with str
            for (var r = 1; r < len; r++) st.characters[start + str.length].remove();
        } else {
            for (var r = 0; r < len; r++) st.characters[start].remove();
        }
    }

    function tokenize(s) {
        // Word runs (letters, digits, apostrophes, accented Latin) or single other chars
        return s.match(/[A-Za-z0-9'\u2019\u00C0-\u024F]+|[\s\S]/g) || [];
    }

    // Word-level diff -> [{start, len, str}] in ascending order, offsets into old text
    function diffOps(a, b) {
        var A = tokenize(a), B = tokenize(b), n = A.length, m = B.length;
        var ops = [];

        if (n * m > 4000000) {
            // Too big for LCS: single prefix/suffix span
            var p = 0;
            while (p < a.length && p < b.length && a.charAt(p) === b.charAt(p)) p++;
            var s = 0;
            while (s < a.length - p && s < b.length - p &&
                   a.charAt(a.length - 1 - s) === b.charAt(b.length - 1 - s)) s++;
            ops.push({ start: p, len: a.length - p - s, str: b.substring(p, b.length - s) });
            return fixInserts(ops, a);
        }

        // LCS table (suffix form)
        var L = [];
        for (var i = 0; i <= n; i++) { L[i] = []; L[i][m] = 0; }
        for (var j = 0; j <= m; j++) L[n][j] = 0;
        for (var i = n - 1; i >= 0; i--) {
            for (var j = m - 1; j >= 0; j--) {
                L[i][j] = (A[i] === B[j]) ? L[i + 1][j + 1] + 1 : Math.max(L[i + 1][j], L[i][j + 1]);
            }
        }

        // Token char offsets in old text
        var off = [], c = 0;
        for (var i = 0; i < n; i++) { off.push(c); c += A[i].length; }
        off.push(c);

        var i = 0, j = 0, cur = null;
        function flush() { if (cur) { ops.push(cur); cur = null; } }
        while (i < n || j < m) {
            if (i < n && j < m && A[i] === B[j]) { flush(); i++; j++; }
            else if (j < m && (i >= n || L[i][j + 1] >= L[i + 1][j])) {
                if (!cur) cur = { start: off[i], len: 0, str: "" };
                cur.str += B[j]; j++;
            } else {
                if (!cur) cur = { start: off[i], len: 0, str: "" };
                cur.len += A[i].length; i++;
            }
        }
        flush();
        return fixInserts(ops, a);
    }

    // Pure insertions (len 0) need an existing char to anchor formatting:
    // absorb the previous char (or next if at start).
    function fixInserts(ops, a) {
        for (var k = 0; k < ops.length; k++) {
            var o = ops[k];
            if (o.len > 0 || !o.str.length || !a.length) continue;
            if (o.start > 0) { o.start -= 1; o.len = 1; o.str = a.charAt(o.start) + o.str; }
            else { o.len = 1; o.str = o.str + a.charAt(0); }
        }
        return ops;
    }
})();
