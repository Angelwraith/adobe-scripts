/*@METADATA{
  "name": "Merge Text",
  "description": "Merges selected text frames top-to-bottom into a single frame, preserving character formatting (kerning, tracking, color, fonts) and exact vertical pixel spacing via per-line leading",
  "version": "1.1",
  "target": "illustrator",
  "tags": ["text", "merge", "combine", "labels"]
}@END_METADATA*/

#target illustrator

// ============================================================================
// Merge Text
// ----------------------------------------------------------------------------
// Customized merge script. Differences from the original community version:
//
//   1. Top-only sort (no separator/sort UI). Stacked labels almost always go
//      top-to-bottom; the dropdown options were redundant for that use case.
//   2. Preserves all character attributes -- kerning, tracking, color, font,
//      manual offsets -- by using textRange.move() instead of overwriting
//      .contents with a separator string (which silently destroys formatting).
//   3. Captures the exact pixel gap between each pair of frames before merging
//      and applies it as leading on the corresponding line of the merged
//      output. Result is visually identical to the original stack.
// ============================================================================

(function () {
    if (app.documents.length === 0) {
        alert("Please open a document first.");
        return;
    }

    var doc = app.activeDocument;
    var sel = doc.selection;

    // Collect text frames from selection
    var tfs = [];
    for (var i = 0; i < sel.length; i++) {
        if (sel[i] && sel[i].typename === "TextFrame") {
            tfs.push(sel[i]);
        }
    }

    if (tfs.length < 2) {
        alert(tfs.length === 1
            ? "Select more than one text frame to merge."
            : "Select two or more text frames to merge.");
        return;
    }

    // Sort top-to-bottom. In Illustrator's coordinate system, a higher 'top'
    // value sits higher on the page, so descending order = visual top first.
    tfs.sort(function (a, b) {
        var diff = b.top - a.top;
        if (diff === 0) return a.left - b.left; // tie-break: left to right
        return diff;
    });

    // Snapshot original top positions BEFORE merging -- geometry shifts as we
    // move text around, and we need the original gaps for leading calculation.
    var tops = [];
    for (var i = 0; i < tfs.length; i++) {
        tops.push(tfs[i].top);
    }

    var target = tfs[0];

    for (var i = 1; i < tfs.length; i++) {
        var source = tfs[i];
        var gap = tops[i - 1] - tops[i]; // exact pixel distance, top-to-top

        // Index where target's current text ends -- the join point
        var oldLength = target.contents.length;

        // Append paragraph break. characters.add() inherits formatting from
        // the previous character, unlike .contents reassignment which resets.
        target.textRange.characters.add("\r");

        // Move source's textRange to the end of target.
        // CRITICAL: move() preserves all character attributes -- kerning,
        // tracking, color, font, size, manual letter offsets. The original
        // script's `arr[i].textRange.contents = separator + arr[i].textRange.contents`
        // was destroying every per-character attribute before the move.
        source.textRange.move(target, ElementPlacement.PLACEATEND);

        // New content lives at indices [oldLength + 1, target.contents.length)
        // (the +1 skips past the \r we just inserted)
        var newStart = oldLength + 1;
        var newEnd = target.contents.length;

        // Find the end of the first line of new content. If the source frame
        // had internal line breaks, anything past the first \r should keep
        // the source's own leading, not our inter-frame gap.
        var firstLineEnd = newEnd;
        for (var c = newStart; c < newEnd; c++) {
            if (target.characters[c].contents === "\r") {
                firstLineEnd = c;
                break;
            }
        }

        // Apply the captured gap as leading. Illustrator uses the MAX leading
        // among a line's characters as the line's effective leading, so we
        // set it on every character of the first new line (and on the \r
        // itself for safety).
        applyLeading(target, oldLength, gap);
        for (var c = newStart; c < firstLineEnd; c++) {
            applyLeading(target, c, gap);
        }

        // Remove the now-empty source frame
        source.remove();
    }

    // Re-select the merged result so the user can see what happened
    doc.selection = [target];

    function applyLeading(frame, charIndex, leadingValue) {
        var ca = frame.characters[charIndex].characterAttributes;
        ca.autoLeading = false;
        ca.leading = leadingValue;
    }
})();
