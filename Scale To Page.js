/*
@METADATA
{
  "name": "Scale To Page",
  "description": "Copies the currently selected 1/10 scale art onto the active artboard at a chosen size, keeping the relative layout, then writes a matching \"Scale 1:N\" label at the bottom of the page. When the page already carries a \"Scale 1:N\" callout (architectural notation like 3/8\" = 1'-0\" works too) the target ratio prefills to match it, so the copies are resized to the page's scale instead of just relabeled. Source art is assumed to be 1:10 and can be overridden. Enter either a percentage OR a target ratio -- the two fields stay in sync as you type. This is a non-blocking palette, so you can pan/zoom the document while it is open; turn on Preview to drop the copies on the page and adjust the size before committing. Keep the source art selected while you work. The label matches the Smart Dimension Tool format so dimensions come out accurate with no extra setup.",
  "version": "1.7",
  "target": "illustrator",
  "tags": ["scale", "copy", "layout", "processor"]
}
@END_METADATA
*/

#target illustrator

(function () {
    'use strict';

    // Close any previous instance of this palette (default engine keeps $.global).
    try {
        if ($.global.__scaleToPagePalette && $.global.__scaleToPagePalette instanceof Window) {
            $.global.__scaleToPagePalette.close();
        }
    } catch (ePrev) {}
    $.global.__scaleToPagePalette = null;

    if (app.documents.length === 0) {
        alert("Please open a document first.");
        return;
    }

    if (!app.activeDocument.selection || app.activeDocument.selection.length === 0) {
        alert("Select the scaled art you want to copy onto the page, then run the script.");
        return;
    }

    var DEFAULT_DENOMINATOR = 10;              // fallback when the page has no scale callout
    var PREVIEW_TAG = "__STP_PREVIEW__";       // name applied to transient preview copies

    // The SOURCE art is what comes out of the extractor -- 1:10 unless the user
    // says otherwise in the palette. A scale callout found on the page is the
    // TARGET: "the page is at 1:16, match it", which is a resize, not a relabel.
    var detectedScale = detectPageScale(app.activeDocument);
    var baseDenominator = DEFAULT_DENOMINATOR;
    var targetDenominator = (detectedScale && detectedScale.ratio > 0) ? detectedScale.ratio : DEFAULT_DENOMINATOR;

    // ------------------------------------------------------------------
    // Number / input helpers (UI-side only -- safe in palette handlers)
    // ------------------------------------------------------------------

    function fmtNum(n) {
        var r = Math.round(n * 100) / 100;
        if (Math.abs(r - Math.round(r)) < 1e-9) return String(Math.round(r));
        return r.toFixed(2).replace(/0+$/, "").replace(/\.$/, "");
    }

    function parsePercent(raw) {
        if (raw === null) return NaN;
        var s = String(raw).replace(/^\s+|\s+$/g, "").toLowerCase();
        if (s.length === 0) return NaN;
        var mult = false;
        if (s.charAt(0) === "x") { mult = true; s = s.substring(1); }
        else if (s.charAt(s.length - 1) === "x") { mult = true; s = s.substring(0, s.length - 1); }
        s = s.replace(/%/g, "").replace(/^\s+|\s+$/g, "");
        var v = parseFloat(s);
        if (isNaN(v)) return NaN;
        return mult ? v * 100 : v;
    }

    function computeFromPct(pct) {
        return { pct: pct, factor: pct / 100, denominator: baseDenominator * 100 / pct };
    }
    function computeFromRatio(denom) {
        var factor = baseDenominator / denom;
        return { pct: factor * 100, factor: factor, denominator: denom };
    }

    // Which field the user touched most recently. When a scale was read off the
    // page the target ratio is the authoritative value, so start there.
    var lastEdited = (detectedScale && detectedScale.ratio > 0) ? "ratio" : "pct";

    function currentParams() {
        var p = null;
        if (lastEdited === "ratio") {
            var r = parseScaleExpression(ratioInput.text);
            if (r) p = computeFromRatio(r.ratio);
        } else {
            var pct = parsePercent(pctInput.text);
            if (!isNaN(pct) && pct > 0) p = computeFromPct(pct);
        }
        if (!p) return null;
        p.scaleString = "1:" + fmtNum(p.denominator);
        return p;
    }

    // ------------------------------------------------------------------
    // Scale detection -- read the drawing scale off the page
    // ------------------------------------------------------------------
    // Looks for a scale callout in the document's text (the "Scale 1:N" label
    // this script and the Smart Dimension Tool write, or architectural notation
    // like 3/8" = 1'-0"). Returns {text, ratio, source} or null.

    function detectPageScale(doc) {
        var frames;
        try {
            frames = doc.textFrames;
        } catch (e) {
            return null;
        }
        if (!frames || frames.length === 0) return null;

        var artRect = null;
        try {
            var abIndex = doc.artboards.getActiveArtboardIndex();
            if (abIndex >= 0 && abIndex < doc.artboards.length) {
                artRect = doc.artboards[abIndex].artboardRect;
            }
        } catch (eAb) {
            artRect = null;
        }

        var limit = frames.length;
        if (limit > 3000) limit = 3000; // sanity cap on very heavy documents

        var best = null;
        var bestScore = -1;

        for (var i = 0; i < limit; i++) {
            var contents;
            try {
                contents = frames[i].contents;
            } catch (eC) {
                continue;
            }
            if (!contents) continue;

            var onArtboard = false;
            if (artRect) {
                try {
                    var b = frames[i].geometricBounds; // [left, top, right, bottom]
                    var cx = (b[0] + b[2]) / 2;
                    var cy = (b[1] + b[3]) / 2;
                    onArtboard = (cx >= artRect[0] && cx <= artRect[2] &&
                                  cy <= artRect[1] && cy >= artRect[3]);
                } catch (eB) {}
            }

            var lines = String(contents).split(/[\r\n]+/);
            for (var j = 0; j < lines.length; j++) {
                var hit = scaleFromLine(lines[j]);
                if (!hit) continue;

                // A line that says "Scale" beats a bare ratio, and a callout on
                // the active artboard beats one elsewhere in the document.
                var score = (hit.labeled ? 2 : 0) + (onArtboard ? 1 : 0);
                if (score > bestScore) {
                    bestScore = score;
                    best = hit;
                }
            }
        }

        return best;
    }

    function scaleFromLine(line) {
        var raw = normalizeScaleText(line).replace(/^\s+|\s+$/g, '');
        if (raw === '') return null;

        var s = raw;
        var labeled = false;

        var labelMatch = s.match(/scale[\s:=\-]+(.+)$/i);
        if (labelMatch) {
            labeled = true;
            s = labelMatch[1];
        }

        s = s.replace(/^[\s\(\[]+/, '').replace(/[\s\)\]\.,;]+$/, '');
        if (s === '') return null;

        var parsed = parseScaleExpression(s);
        if (!parsed) return null;

        parsed.labeled = labeled;
        parsed.source = raw;
        return parsed;
    }

    // Parse a scale expression into {text: "A:B", ratio: N} where N is what an
    // on-page measurement is multiplied by to get real-world size.
    //   "1:10"           -> ratio 10
    //   "2:1"            -> ratio 0.5
    //   "3/8\" = 1'-0\"" -> ratio 32  (text "1:32")
    function parseScaleExpression(text) {
        var s = normalizeScaleText(text).replace(/^\s+|\s+$/g, '');
        if (s === '') return null;

        var ratioMatch = s.match(/^(\d+(?:\.\d+)?)\s*:\s*(\d+(?:\.\d+)?)$/);
        if (ratioMatch) {
            var a = parseFloat(ratioMatch[1]);
            var b = parseFloat(ratioMatch[2]);
            if (isNaN(a) || isNaN(b) || a <= 0 || b <= 0) return null;
            return { text: fmtNum(a) + ":" + fmtNum(b), ratio: b / a };
        }

        var eqIdx = s.indexOf('=');
        if (eqIdx > 0 && eqIdx < s.length - 1) {
            var leftIn = parseAsInches(s.substring(0, eqIdx));
            var rightIn = parseAsInches(s.substring(eqIdx + 1));
            if (leftIn === null || rightIn === null || leftIn <= 0 || rightIn <= 0) {
                return null;
            }
            var r = rightIn / leftIn;
            return { text: "1:" + fmtNum(r), ratio: r };
        }

        return null;
    }

    // Parse an expression representing a length in inches. Handles feet-inches
    // ("1'-6\""), feet only ("1'"), fractions ("3/8"), mixed numbers ("1 1/2"),
    // decimals, and whole numbers.
    function parseAsInches(text) {
        if (text === undefined || text === null) return null;
        var s = normalizeScaleText(text).replace(/^\s+|\s+$/g, '');
        if (s === '') return null;

        var totalInches = 0;

        var feetMatch = s.match(/^(\d+(?:\.\d+)?)\s*'/);
        if (feetMatch) {
            totalInches += parseFloat(feetMatch[1]) * 12;
            s = s.substring(feetMatch[0].length);
            s = s.replace(/^[\s\-]+/, '');
        }

        s = s.replace(/"\s*$/, '').replace(/^\s+|\s+$/g, '');

        if (s === '') return totalInches;

        var mixed = s.match(/^(\d+)[\s\-]+(\d+)\s*\/\s*(\d+)$/);
        if (mixed) {
            var mDenom = parseFloat(mixed[3]);
            if (mDenom === 0) return null;
            totalInches += parseFloat(mixed[1]) + (parseFloat(mixed[2]) / mDenom);
            return totalInches;
        }

        var frac = s.match(/^(\d+)\s*\/\s*(\d+)$/);
        if (frac) {
            var fDenom = parseFloat(frac[2]);
            if (fDenom === 0) return null;
            totalInches += parseFloat(frac[1]) / fDenom;
            return totalInches;
        }

        if (/^\d+(?:\.\d+)?$/.test(s)) {
            totalInches += parseFloat(s);
            return totalInches;
        }

        return null;
    }

    // Normalize the typographic characters real proofs contain (prime and
    // double-prime marks, curly quotes, en/em dashes, non-breaking spaces,
    // Unicode fraction glyphs) down to plain ASCII. Done by char code so this
    // source file stays pure ASCII.
    function normalizeScaleText(input) {
        if (input === undefined || input === null) return '';
        var src = String(input);
        var out = [];
        for (var i = 0; i < src.length; i++) {
            var c = src.charCodeAt(i);
            if (c === 0x201C || c === 0x201D || c === 0x201E || c === 0x201F ||
                c === 0x2033 || c === 0x3003 || c === 0x301D || c === 0x301E || c === 0x301F) {
                out.push('"');   // double-prime / smart double quotes -> "
            } else if (c === 0x2018 || c === 0x2019 || c === 0x201A || c === 0x201B ||
                       c === 0x2032 || c === 0x00B4 || c === 0x0060) {
                out.push("'");   // prime / smart single quotes / acute / backtick -> '
            } else if (c === 0x2010 || c === 0x2011 || c === 0x2012 || c === 0x2013 ||
                       c === 0x2014 || c === 0x2015 || c === 0x2212) {
                out.push('-');   // assorted dashes / minus -> hyphen
            } else if (c === 0x00A0 || (c >= 0x2000 && c <= 0x200B) ||
                       c === 0x202F || c === 0x205F || c === 0x3000) {
                out.push(' ');   // assorted spaces -> space
            } else {
                var fr = fractionForCode(c);
                out.push(fr !== null ? (' ' + fr + ' ') : src.charAt(i));
            }
        }
        return out.join('');
    }

    function fractionForCode(c) {
        switch (c) {
            case 0x00BC: return '1/4';
            case 0x00BD: return '1/2';
            case 0x00BE: return '3/4';
            case 0x2153: return '1/3';
            case 0x2154: return '2/3';
            case 0x2155: return '1/5';
            case 0x2156: return '2/5';
            case 0x2157: return '3/5';
            case 0x2158: return '4/5';
            case 0x2159: return '1/6';
            case 0x215A: return '5/6';
            case 0x215B: return '1/8';
            case 0x215C: return '3/8';
            case 0x215D: return '5/8';
            case 0x215E: return '7/8';
        }
        return null;
    }

    function truncateForDisplay(text, maxLen) {
        var s = String(text || '').replace(/^\s+|\s+$/g, '');
        if (s.length <= maxLen) return s;
        return s.substring(0, maxLen - 3) + '...';
    }

    // ------------------------------------------------------------------
    // Document work runs in Illustrator's MAIN engine via BridgeTalk.
    // (A palette's own handlers can't reliably touch app.activeDocument --
    //  doing so directly is what froze the earlier version.)
    // ------------------------------------------------------------------

    function runInMain(body, onDone) {
        try {
            var bt = new BridgeTalk();
            bt.target = "illustrator";
            bt.body = body;
            bt.onResult = function (res) { if (onDone) onDone(res.body); };
            bt.onError = function (err) { if (onDone) onDone("ERR:" + ((err && err.body) || "unknown")); };
            bt.send();
        } catch (e) {
            if (onDone) onDone("ERR:" + e.message);
        }
    }

    // Build the placement script (as a string) for the main engine.
    //   isPreview true  -> tag copies with PREVIEW_TAG, keep source selected, leave existing labels alone
    //   isPreview false -> remove old "Scale" labels, add the real label, select the copies
    // Existing PREVIEW_TAG items are always cleared first (so re-preview / commit don't stack).
    // Preview copies live on their own layer named PREVIEW_TAG. Creating/removing
    // a layer is instant -- unlike scanning d.pageItems, which is brutally slow on
    // large production files (that scan was the source of the lag).
    function buildPlaceScript(factor, scaleString, doLabel, isPreview) {
        var s = "";
        s += "var __r='OK';";
        s += "try{";
        s += "if(app.documents.length===0){__r='NODOC';}else{";
        s += "var d=app.activeDocument;";
        s += "var sel=d.selection;";
        s += "if(!sel||sel.length===0){__r='NOSEL';}else{";
        s += "var LYR='" + PREVIEW_TAG + "';";
        s += "var factor=" + factor + ";";
        s += "function cb(items){var L=1/0,T=-1/0,R=-1/0,B=1/0;for(var k=0;k<items.length;k++){var b=items[k].geometricBounds;if(b[0]<L)L=b[0];if(b[1]>T)T=b[1];if(b[2]>R)R=b[2];if(b[3]<B)B=b[3];}return [L,T,R,B];}";
        // Capture the source, then remove any existing preview layer (fast).
        s += "var src=[];for(var i=0;i<sel.length;i++){src.push(sel[i]);}";
        s += "try{for(var li=d.layers.length-1;li>=0;li--){if(d.layers[li].name===LYR){d.layers[li].remove();}}}catch(e){}";
        // Duplicate + scale as a unit about the collective top-left.
        s += "var copies=[];for(var s2=0;s2<src.length;s2++){copies.push(src[s2].duplicate());}";
        s += "var pre=cb(copies);var ax=pre[0],ay=pre[1];";
        s += "for(var c=0;c<copies.length;c++){var it=copies[c];var bb=it.geometricBounds;var oL=bb[0],oT=bb[1];";
        s += "it.resize(factor*100,factor*100,true,true,true,true,true,Transformation.TOPLEFT);";
        s += "var nL=ax+factor*(oL-ax);var nT=ay+factor*(oT-ay);it.translate(nL-oL,nT-oT);}";
        // Center on the active artboard.
        s += "var ai=d.artboards.getActiveArtboardIndex();var ar=d.artboards[ai].artboardRect;";
        s += "var acx=(ar[0]+ar[2])/2,acy=(ar[1]+ar[3])/2;";
        s += "var post=cb(copies);var ccx=(post[0]+post[2])/2,ccy=(post[1]+post[3])/2;var dx=acx-ccx,dy=acy-ccy;";
        s += "for(var m=0;m<copies.length;m++){copies[m].translate(dx,dy);}";
        s += "var lbl=null;";
        if (doLabel) {
            if (!isPreview) {
                // Remove existing "Scale ..." labels sitting on the active artboard.
                s += "try{for(var t=d.textFrames.length-1;t>=0;t--){var tf=d.textFrames[t];var cn='';try{cn=tf.contents;}catch(e){cn='';}";
                s += "if(cn.toLowerCase().indexOf('scale ')!==0)continue;var tb=tf.geometricBounds;var lcx=(tb[0]+tb[2])/2,lcy=(tb[1]+tb[3])/2;";
                s += "if(lcx>=ar[0]&&lcx<=ar[2]&&lcy<=ar[1]&&lcy>=ar[3])tf.remove();}}catch(e){}";
            }
            s += "var lay=d.activeLayer;try{if(lay.locked||!lay.visible){for(var Lz=0;Lz<d.layers.length;Lz++){if(!d.layers[Lz].locked&&d.layers[Lz].visible){lay=d.layers[Lz];break;}}}}catch(e){}";
            s += "try{var lx=ar[0]+(4*72);var ly=ar[1]-(6.9816*72);lbl=lay.textFrames.add();lbl.contents='Scale " + scaleString + "';";
            s += "var cs=null;try{for(var ci=0;ci<d.characterStyles.length;ci++){if(d.characterStyles[ci].name==='DimStyle'){cs=d.characterStyles[ci];break;}}}catch(e){}";
            s += "if(cs){lbl.textRange.characterAttributes.characterStyle=cs;}";
            s += "lbl.textRange.characterAttributes.size=7;lbl.textRange.paragraphAttributes.justification=Justification.CENTER;";
            s += "lbl.top=ly;lbl.left=lx-(lbl.width/2);}catch(e){}";
        }
        if (isPreview) {
            // Move the copies (and label) onto a fresh preview layer; keep the
            // source selected so re-preview / commit still see it.
            s += "try{var pl=d.layers.add();pl.name=LYR;for(var mv=copies.length-1;mv>=0;mv--){try{copies[mv].moveToBeginning(pl);}catch(e){}}";
            s += "if(lbl){try{lbl.moveToBeginning(pl);}catch(e){}}}catch(e){}";
            s += "try{d.selection=src;}catch(e){}";
        } else {
            s += "try{d.selection=null;for(var sc=0;sc<copies.length;sc++){copies[sc].selected=true;}}catch(e){}";
        }
        s += "app.redraw();";
        s += "}}";
        s += "}catch(e){__r='ERR:'+e.message;}";
        s += "__r;";
        return s;
    }

    // Remove the preview layer (and everything on it) -- instant, no page scan.
    function buildClearScript() {
        var s = "";
        s += "var __r='OK';try{if(app.documents.length>0){var d=app.activeDocument;";
        s += "for(var li=d.layers.length-1;li>=0;li--){try{if(d.layers[li].name==='" + PREVIEW_TAG + "')d.layers[li].remove();}catch(e){}}";
        s += "app.redraw();}}catch(e){__r='ERR:'+e.message;}__r;";
        return s;
    }

    var busy = false; // guard against overlapping BridgeTalk operations

    function reportProblem(code) {
        if (code === "NOSEL") {
            statusText.text = "Select the source art first (it must stay selected).";
        } else if (code === "NODOC") {
            statusText.text = "No document open.";
        } else if (code && code.indexOf("ERR:") === 0) {
            statusText.text = "Error: " + code.substring(4);
        } else {
            statusText.text = "";
        }
    }

    function renderPreview() {
        if (busy) return;
        var p = currentParams();
        if (!p) return;
        busy = true;
        runInMain(buildPlaceScript(p.factor, p.scaleString, cbLabel.value, true), function (r) {
            busy = false;
            reportProblem(r === "OK" ? "" : r);
        });
    }

    function clearPreview(onDone) {
        runInMain(buildClearScript(), function (r) { if (onDone) onDone(r); });
    }

    // ------------------------------------------------------------------
    // Palette UI (non-blocking)
    // ------------------------------------------------------------------

    var dlg = new Window("palette", "Scale To Page");
    dlg.orientation = "column";
    dlg.alignChildren = "fill";
    dlg.margins = 16;
    dlg.spacing = 10;
    $.global.__scaleToPagePalette = dlg;

    var srcCount = app.activeDocument.selection.length;
    dlg.add("statictext", undefined, "Selected art: " + srcCount + " object(s)");

    var baseGroup = dlg.add("group");
    var baseLbl = baseGroup.add("statictext", undefined, "Source art:");
    baseLbl.preferredSize.width = 80;
    var baseInput = baseGroup.add("edittext", undefined, "1:" + fmtNum(baseDenominator));
    baseInput.characters = 8;
    var baseNote = baseGroup.add("statictext", undefined, "scale of the art you selected");
    baseNote.preferredSize.width = 250;

    var pctGroup = dlg.add("group");
    var pctLbl = pctGroup.add("statictext", undefined, "Resize to:");
    pctLbl.preferredSize.width = 80;
    var pctInput = pctGroup.add("edittext", undefined, fmtNum(baseDenominator * 100 / targetDenominator) + "%");
    pctInput.characters = 8;
    pctGroup.add("statictext", undefined, "% of current size");

    var ratioGroup = dlg.add("group");
    var ratioLbl = ratioGroup.add("statictext", undefined, "Target ratio:");
    ratioLbl.preferredSize.width = 80;
    var ratioInput = ratioGroup.add("edittext", undefined, "1:" + fmtNum(targetDenominator));
    ratioInput.characters = 8;
    var ratioNote = ratioGroup.add("statictext", undefined, "");
    ratioNote.preferredSize.width = 250;
    ratioNote.text = detectedScale
        ? ('matching page: "' + truncateForDisplay(detectedScale.source, 28) + '"')
        : "e.g. 1:12   (updates with the percentage)";

    var previewText = dlg.add("statictext", undefined, "New scale on page:  1:10   (100% of current)");
    previewText.preferredSize.width = 360;
    previewText.graphics.font = ScriptUI.newFont(previewText.graphics.font.name, "BOLD", 13);

    var optPanel = dlg.add("panel", undefined, "Options");
    optPanel.orientation = "column";
    optPanel.alignChildren = "left";
    optPanel.margins = 12;
    optPanel.add("statictext", undefined, "Copies are always placed on the ACTIVE artboard.");
    var cbLabel = optPanel.add("checkbox", undefined, "Add \"Scale 1:N\" label (matches Smart Dimension Tool exactly)");
    cbLabel.value = true;
    var cbPreview = optPanel.add("checkbox", undefined, "Preview on page (keeps this palette open so you can adjust)");
    cbPreview.value = false;

    var statusText = dlg.add("statictext", undefined, "");
    statusText.preferredSize.width = 360;

    // ------------------------------------------------------------------
    // Field sync + live text preview
    // ------------------------------------------------------------------
    var syncing = false;

    function refreshPreviewText() {
        var p = currentParams();
        if (!p) { previewText.text = "New scale on page:  --"; return; }
        var txt = "New scale on page:  1:" + fmtNum(p.denominator) + "   (" + fmtNum(p.pct) + "% of current)";
        if (Math.abs(p.denominator - Math.round(p.denominator)) > 1e-9) txt += "   <-- has decimals";
        previewText.text = txt;
    }

    function syncFromPct() {
        if (syncing) return;
        syncing = true;
        lastEdited = "pct";
        var pct = parsePercent(pctInput.text);
        if (!isNaN(pct) && pct > 0) ratioInput.text = "1:" + fmtNum(baseDenominator * 100 / pct);
        syncing = false;
        refreshPreviewText();
    }
    function syncFromRatio() {
        if (syncing) return;
        syncing = true;
        lastEdited = "ratio";
        var r = parseScaleExpression(ratioInput.text);
        if (r) pctInput.text = fmtNum(baseDenominator * 100 / r.ratio) + "%";
        syncing = false;
        refreshPreviewText();
    }
    // Editing the source art's scale keeps whichever field was last touched and
    // recalculates the other one against the new base.
    function syncFromBase() {
        if (syncing) return;
        var parsed = parseScaleExpression(baseInput.text);
        if (!parsed || !(parsed.ratio > 0)) return;
        baseDenominator = parsed.ratio;
        syncing = true;
        if (lastEdited === "ratio") {
            var r = parseScaleExpression(ratioInput.text);
            if (r) pctInput.text = fmtNum(baseDenominator * 100 / r.ratio) + "%";
        } else {
            var pct = parsePercent(pctInput.text);
            if (!isNaN(pct) && pct > 0) ratioInput.text = "1:" + fmtNum(baseDenominator * 100 / pct);
        }
        syncing = false;
        refreshPreviewText();
    }
    pctInput.onChanging = syncFromPct;
    ratioInput.onChanging = syncFromRatio;
    baseInput.onChanging = syncFromBase;

    // Refresh the on-page preview when a field edit is committed or options change.
    function onFieldCommit() { if (cbPreview.value) renderPreview(); }
    pctInput.onChange = onFieldCommit;
    ratioInput.onChange = onFieldCommit;
    baseInput.onChange = onFieldCommit;

    cbPreview.onClick = function () {
        if (cbPreview.value) renderPreview();
        else clearPreview();
    };
    cbLabel.onClick = function () { if (cbPreview.value) renderPreview(); };

    refreshPreviewText();

    // ------------------------------------------------------------------
    // Buttons
    // ------------------------------------------------------------------
    var btnGroup = dlg.add("group");
    btnGroup.alignment = "right";
    var cancelBtn = btnGroup.add("button", undefined, "Cancel", { name: "cancel" });
    var okBtn = btnGroup.add("button", undefined, "OK");
    dlg.defaultElement = okBtn;
    dlg.cancelElement = cancelBtn;

    var committed = false;

    okBtn.onClick = function () {
        if (busy) return;
        var p = currentParams();
        if (!p) {
            alert("Please enter a valid resize amount.\n\nExamples:\n  50%    (half of current size)\n  2x     (double)\n  1:12   (target on-page scale)");
            return;
        }
        busy = true;
        statusText.text = "Placing...";
        runInMain(buildPlaceScript(p.factor, p.scaleString, cbLabel.value, false), function (r) {
            busy = false;
            if (r === "OK") {
                committed = true;
                dlg.close();
            } else {
                reportProblem(r);
            }
        });
    };

    cancelBtn.onClick = function () { dlg.close(); };

    // Closing any non-commit way removes the transient preview from the page.
    dlg.onClose = function () {
        if (!committed) clearPreview();
        $.global.__scaleToPagePalette = null;
        return true;
    };

    dlg.center();
    dlg.show(); // palette -> non-blocking; the document stays interactive
})();
