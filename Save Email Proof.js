/*
@METADATA
{
  "name": "Save email proof",
  "description": "Save an EmProof PDF (downsampled); also save a full-res Proof if the doc has unsaved changes",
  "version": "2.3",
  "target": "illustrator",
  "tags": ["proof", "email", "pdf", "save", "export"]
}
@END_METADATA
*/

/*
  Settings are hardcoded from the EmProof.joboptions and IllustratorDefault.joboptions
  presets so this script never depends on those presets being installed. Verified
  against Illustrator 2026 (30.6): the two presets differ only in the size-driving
  options below.

    EmProof -> downsample images to 150ppi (color/gray) / 300ppi (mono), Average
               method; no editability, no Acrobat layers, no thumbnails.  (~4 MB)
    Proof   -> full resolution (no downsampling); preserve editability, Acrobat
               layers + thumbnails on.  (full size)

  Behavior: the EmProof is always saved. The full-res Proof is ALSO saved whenever
  the document has unsaved changes, so those changes are preserved in a lossless,
  editable PDF and are not lost by only saving the downsampled EmProof.

  Shared: Acrobat 1.6, ZIP/Flate image compression (color & gray), compress art,
  embed fonts (subset 100%), colors unchanged, fast-web-view off.

  NOTE: monochromeCompression is intentionally NOT set -- Illustrator 2026 rejects
  the ZIP/CCITT enum values for that property, and the default is fine (mono images
  have negligible size impact). set() also guards any value a future version rejects.
*/

(function () {
    var EMPROOF_SUFFIX = "_EmProof";
    var PROOF_SUFFIX = "_Proof";

    // --- Preflight ---
    if (app.documents.length === 0) {
        alert("No document is open.");
        return;
    }

    var doc = app.activeDocument;
    var sep = ($.os.indexOf("Windows") !== -1) ? "\\" : "/";

    // Capture BEFORE any saveAs (saveAs marks the document as saved).
    // doc.saved === false means there are unsaved changes (or it was never saved).
    var hasUnsavedChanges = (doc.saved === false);

    // Tolerant setter: skips (and records) any property/enum this AI version rejects.
    var rejected = [];
    function set(o, prop, val) {
        try { o[prop] = val; } catch (e) { rejected.push(prop); }
    }

    // Determine the folder to save into (folder containing the document).
    var saveFolder;
    try {
        saveFolder = doc.path; // throws if the document has never been saved
        if (!saveFolder || !saveFolder.exists) throw new Error("no path");
    } catch (e) {
        var picked = Folder.selectDialog("This document hasn't been saved yet. Choose a folder to save the proofs into:");
        if (!picked) return; // user cancelled
        saveFolder = picked;
    }

    // Build the base name: strip extension and any existing proof suffix.
    var baseName = decodeURI(doc.name).replace(/\.[^\.]+$/, "");
    baseName = baseName.replace(/(_EmProof|_Proof)$/i, "");

    // --- Settings shared by both presets ---
    function applyShared(o) {
        set(o, "compatibility", PDFCompatibility.ACROBAT7);   // PDF 1.6
        set(o, "optimization", false);                        // Optimize (fast web view) off
        set(o, "compressArt", true);                          // compress line art & text
        set(o, "fontSubsetThreshold", 100);                   // MaxSubsetPct 100
        set(o, "colorConversionID", ColorConversion.None);    // LeaveColorUnchanged
        set(o, "viewAfterSaving", false);
        set(o, "colorCompression", CompressionQuality.ZIP8BIT);      // Flate/ZIP
        set(o, "grayscaleCompression", CompressionQuality.ZIP8BIT);  // Flate/ZIP
        // monochromeCompression left at default (see header note).
    }

    // --- EmProof preset (downsampled, smaller file) ---
    function makeEmProofOptions() {
        var o = new PDFSaveOptions();
        applyShared(o);
        set(o, "preserveEditability", false);   // PreserveEditing false
        set(o, "acrobatLayers", false);         // IncludeLayers false
        set(o, "generateThumbnails", false);    // DoThumbnails false

        set(o, "colorDownsamplingMethod", DownsampleMethod.AVERAGEDOWNSAMPLE);
        set(o, "colorDownsampling", 150);
        set(o, "colorDownsamplingImageThreshold", 1.5);
        set(o, "grayscaleDownsamplingMethod", DownsampleMethod.AVERAGEDOWNSAMPLE);
        set(o, "grayscaleDownsampling", 150);
        set(o, "grayscaleDownsamplingImageThreshold", 1.5);
        set(o, "monochromeDownsamplingMethod", DownsampleMethod.AVERAGEDOWNSAMPLE);
        set(o, "monochromeDownsampling", 300);
        set(o, "monochromeDownsamplingImageThreshold", 1.5);
        return o;
    }

    // --- Proof preset (Illustrator Default: full resolution, no downsampling) ---
    function makeProofOptions() {
        var o = new PDFSaveOptions();
        applyShared(o);
        set(o, "preserveEditability", true);    // PreserveEditing true
        set(o, "acrobatLayers", true);          // IncludeLayers true
        set(o, "generateThumbnails", true);     // DoThumbnails true
        // No downsampling properties set -> Illustrator keeps full resolution.
        return o;
    }

    function savePDF(suffix, options) {
        var outFile = new File(saveFolder.fsName + sep + baseName + suffix + ".pdf");
        doc.saveAs(outFile, options);
        return outFile;
    }

    try {
        // If the document has unsaved changes, save the full-res Proof first so a
        // lossless, editable copy of the current state is preserved.
        if (hasUnsavedChanges) {
            savePDF(PROOF_SUFFIX, makeProofOptions());
        }

        // Always save the EmProof. Saved from the same in-memory artwork, so it is
        // NOT built from any earlier save.
        savePDF(EMPROOF_SUFFIX, makeEmProofOptions());

        // Only surface a popup if a setting was rejected (should not happen).
        if (rejected.length) {
            var seen = {}, uniq = [];
            for (var i = 0; i < rejected.length; i++) {
                if (!seen[rejected[i]]) { seen[rejected[i]] = true; uniq.push(rejected[i]); }
            }
            alert("Some settings were rejected (tell Claude):\n" + uniq.join(", "));
        }
    } catch (err) {
        alert("Error saving proof.\n\n" + err.message);
    }
})();
