/*
@METADATA
{
  "name": "Scale To Page",
  "description": "Copies the currently selected 1/10 scale art onto the active artboard at a chosen size, keeping the relative layout, then writes a matching \"Scale 1:N\" label at the bottom of the page. Starting scale is assumed to be 1:10, so 50% -> 1:20, 200% -> 1:5, etc. Enter either a percentage OR a target ratio -- the two fields stay in sync as you type. This is a non-blocking palette, so you can pan/zoom the document while it is open; turn on Preview to drop the copies on the page and adjust the size before committing. Keep the source art selected while you work. The label matches the Smart Dimension Tool format so dimensions come out accurate with no extra setup.",
  "version": "1.5",
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
        alert("Select the 1/10 scale art you want to copy onto the page, then run the script.");
        return;
    }

    var BASE_DENOMINATOR = 10;                 // art is always extracted at 1:10 scale
    var PREVIEW_TAG = "__STP_PREVIEW__";       // name applied to transient preview copies

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

    function parseRatio(raw) {
        if (raw === null) return null;
        var s = String(raw).replace(/^\s+|\s+$/g, "");
        if (s.indexOf(":") === -1) return null;
        var parts = s.split(":");
        var a = parseFloat(parts[0]);
        var b = parseFloat(parts[1]);
        if (isNaN(a) || isNaN(b) || a <= 0 || b <= 0) return null;
        return { a: a, b: b, denominator: b / a };
    }

    function computeFromPct(pct) {
        return { pct: pct, factor: pct / 100, denominator: BASE_DENOMINATOR * 100 / pct };
    }
    function computeFromRatio(denom) {
        var factor = BASE_DENOMINATOR / denom;
        return { pct: factor * 100, factor: factor, denominator: denom };
    }

    var lastEdited = "pct"; // which field the user touched most recently

    function currentParams() {
        var p = null;
        if (lastEdited === "ratio") {
            var r = parseRatio(ratioInput.text);
            if (r) p = computeFromRatio(r.denominator);
        } else {
            var pct = parsePercent(pctInput.text);
            if (!isNaN(pct) && pct > 0) p = computeFromPct(pct);
        }
        if (!p) return null;
        p.scaleString = "1:" + fmtNum(p.denominator);
        return p;
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
    dlg.add("statictext", undefined, "Selected art: " + srcCount + " object(s)   (starting scale 1:10)");

    var pctGroup = dlg.add("group");
    pctGroup.add("statictext", undefined, "Resize to:");
    var pctInput = pctGroup.add("edittext", undefined, "100%");
    pctInput.characters = 8;
    pctGroup.add("statictext", undefined, "% of current size");

    var ratioGroup = dlg.add("group");
    var ratioLbl = ratioGroup.add("statictext", undefined, "Target ratio:");
    ratioLbl.preferredSize.width = 62;
    var ratioInput = ratioGroup.add("edittext", undefined, "1:10");
    ratioInput.characters = 8;
    ratioGroup.add("statictext", undefined, "e.g. 1:12   (updates with the percentage)");

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
        if (!isNaN(pct) && pct > 0) ratioInput.text = "1:" + fmtNum(BASE_DENOMINATOR * 100 / pct);
        syncing = false;
        refreshPreviewText();
    }
    function syncFromRatio() {
        if (syncing) return;
        syncing = true;
        lastEdited = "ratio";
        var r = parseRatio(ratioInput.text);
        if (r) pctInput.text = fmtNum(BASE_DENOMINATOR * 100 / r.denominator) + "%";
        syncing = false;
        refreshPreviewText();
    }
    pctInput.onChanging = syncFromPct;
    ratioInput.onChanging = syncFromRatio;

    // Refresh the on-page preview when a field edit is committed or options change.
    function onFieldCommit() { if (cbPreview.value) renderPreview(); }
    pctInput.onChange = onFieldCommit;
    ratioInput.onChange = onFieldCommit;

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
