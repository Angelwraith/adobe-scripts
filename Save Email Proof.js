/*
@METADATA
{
  "name": "Save email proof",
  "description": "Save a downsampled Em proof then a full-res proof (proof saved last so the doc keeps the proof name)",
  "version": "2.4",
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

    Em proof -> downsample images to 150ppi (color/gray) / 300ppi (mono), Average
                method; no editability, no Acrobat layers, no thumbnails.  (~4 MB)
    Proof    -> full resolution (no downsampling); preserve editability, Acrobat
                layers + thumbnails on.  (full size)

  Order: the Em proof is saved FIRST and the full-res proof SECOND, so after the
  script runs the active document is the full-res proof (not the Em version) and
  you can keep working under the normal proof name. Saving the full-res proof last
  also preserves any unsaved changes in a lossless, editable PDF.

  Naming: the proof-type token is detected from the current file name so both
  conventions work --
      *_Proof / *.ai / etc.  -> saves  *_EmProof  and  *_Proof
      *_ProdProof            -> saves  *_EmProdProof  and  *_ProdProof
  (Any existing _Em/_Proof/_ProdProof suffix on the open file is stripped first.)

  Shared: Acrobat 1.6, ZIP/Flate image compression (color & gray), compress art,
  embed fonts (subset 100%), colors unchanged, fast-web-view off.

  NOTE: monochromeCompression is intentionally NOT set -- Illustrator 2026 rejects
  the ZIP/CCITT enum values for that property, and the default is fine (mono images
  have negligible size impact). set() also guards any value a future version rejects.
*/

(function () {
    // --- Preflight ---
    if (app.documents.length === 0) {
        alert("No document is open.");
        return;
    }

    var doc = app.activeDocument;
    var sep = ($.os.indexOf("Windows") !== -1) ? "\\" : "/";

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

    // --- Base name + proof-naming style, detected from the current file name ---
    // Recognizes a trailing _Proof / _EmProof (standard) or _ProdProof / _EmProdProof.
    var rawName = decodeURI(doc.name).replace(/\.[^\.]+$/, "");
    var baseName = rawName;
    var isProd = false;
    var m = rawName.match(/^(.*)_(Em)?(Prod)?Proof$/i);
    if (m) {
        baseName = m[1];
        isProd = (m[3] !== undefined && m[3] !== null && m[3] !== "");
    }
    var emSuffix    = isProd ? "_EmProdProof" : "_EmProof";
    var proofSuffix = isProd ? "_ProdProof"   : "_Proof";

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

    // --- Em proof preset (downsampled, smaller file) ---
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
        // 1) Em proof FIRST (downsampled). Built from the current in-memory artwork.
        savePDF(emSuffix, makeEmProofOptions());

        // 2) Full-res proof LAST, so the active document ends up as the proof file
        //    (normal name) and you can keep working. Also built from the same
        //    in-memory artwork, so it is full resolution (not the downsampled Em).
        savePDF(proofSuffix, makeProofOptions());

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
