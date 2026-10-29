/*
@METADATA
{
  "name": "Bit-Knife Cut Converter",
  "description": "Converts selected bit (router) cut paths to knife cuts (split at sharp corners, 0.5in overshoots into waste only, colinear edges merged) or knife cuts back to closed bit paths",
  "version": "1.1",
  "target": "illustrator",
  "tags": ["Cut", "Knife", "Router", "Bit", "processors"]
}
@END_METADATA
*/

#target illustrator

// ============================================================================
// SETTINGS
// ============================================================================
var BKC_OVERSHOOT_IN = 0.5;       // Overshoot past convex corners
var BKC_CORNER_DEG = 30;          // Direction change that counts as a sharp corner
var BKC_MARGIN_IN = 0.125;        // Shortened overshoots stop this far before another cut
var BKC_MIN_EXT_IN = 0.03;        // Overshoots shorter than this are dropped
var BKC_STROKE_PT = 5;            // Output cut line weight (pt, scaled for Large Canvas)
var BKC_TOL_IN = 0.01;            // Geometric tolerance
var BKC_TRIM_WINDOW_FACTOR = 1.5; // Knife->bit: stubs up to OVERSHOOT * this are trimmed
var BKC_INNER_CUTS_AT_BOTTOM = true; // true = inner cuts stacked below outer cuts (cutter runs bottom-up)

var BKC_COLOR_OUTSIDE = "CutThrough2-Outside";
var BKC_COLOR_INSIDE = "CutThrough1-Inside";
var BKC_COLOR_KNIFE = "CutThrough-Knifecut";
var BKC_OBSTACLE_LAYERS = ["CutThrough2-Outside", "CutThrough1-Inside", "PinMountHoles-Inside", "CutThrough-Knifecut", "CutContour"];

// ============================================================================
// MAIN
// ============================================================================
(function () {
    try {
        bkcMain();
    } catch (e) {
        alert("Bit-Knife Cut Converter error:\n" + e.message + (e.line ? "\nLine: " + e.line : ""));
    }
})();

function bkcMain() {
    if (app.documents.length === 0) {
        alert("Please open a document first.");
        return;
    }
    var doc = app.activeDocument;
    var sel = doc.selection;
    if (!sel || sel.length === 0) {
        alert("Select the cut paths to convert first.");
        return;
    }

    var sf = 1;
    try { if (doc.scaleFactor) { sf = doc.scaleFactor; } } catch (eSf) {}
    var IN = 72 / sf;
    var OV = BKC_OVERSHOOT_IN * IN;
    var MARGIN = BKC_MARGIN_IN * IN;
    var MIN_EXT = BKC_MIN_EXT_IN * IN;
    var TOL = BKC_TOL_IN * IN;
    var CORNER_RAD = BKC_CORNER_DEG * Math.PI / 180;
    var PROBE = 0.02 * IN;
    var ANG_EPS = Math.sin(0.05 * Math.PI / 180);

    // ------------------------------------------------------------------
    // Vector + bezier helpers
    // ------------------------------------------------------------------
    function cp(p) { return [p[0], p[1]]; }
    function sub(a, b) { return [a[0] - b[0], a[1] - b[1]]; }
    function add(a, b) { return [a[0] + b[0], a[1] + b[1]]; }
    function mul(a, s) { return [a[0] * s, a[1] * s]; }
    function dot(a, b) { return a[0] * b[0] + a[1] * b[1]; }
    function cross(a, b) { return a[0] * b[1] - a[1] * b[0]; }
    function vlen(a) { return Math.sqrt(a[0] * a[0] + a[1] * a[1]); }
    function dist(a, b) { return vlen(sub(a, b)); }
    function norm(a) { var l = vlen(a); return l < 1e-12 ? [0, 0] : [a[0] / l, a[1] / l]; }
    function lerp(a, b, t) { return [a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t]; }
    function clamp(v, lo, hi) { return v < lo ? lo : (v > hi ? hi : v); }

    function bez(c, t) {
        var mt = 1 - t;
        var a = mt * mt * mt, b = 3 * mt * mt * t, d = 3 * mt * t * t, e = t * t * t;
        return [a * c[0][0] + b * c[1][0] + d * c[2][0] + e * c[3][0],
                a * c[0][1] + b * c[1][1] + d * c[2][1] + e * c[3][1]];
    }
    function splitCubic(c, t) {
        var p01 = lerp(c[0], c[1], t), p12 = lerp(c[1], c[2], t), p23 = lerp(c[2], c[3], t);
        var p012 = lerp(p01, p12, t), p123 = lerp(p12, p23, t);
        var m = lerp(p012, p123, t);
        return [[cp(c[0]), p01, p012, m], [cp(m), p123, p23, cp(c[3])]];
    }
    function subCubic(c, t0, t1) {
        var r = c;
        if (t1 < 1 - 1e-12) { r = splitCubic(r, t1)[0]; }
        if (t0 > 1e-12) { r = splitCubic(r, t1 > 1e-12 ? t0 / t1 : 0)[1]; }
        return [cp(r[0]), cp(r[1]), cp(r[2]), cp(r[3])];
    }
    function isLinear(c) {
        var d = dist(c[0], c[3]);
        if (d < 1e-9) { return dist(c[0], c[1]) < 1e-9 && dist(c[2], c[3]) < 1e-9; }
        var u = norm(sub(c[3], c[0]));
        var hs = [c[1], c[2]];
        for (var i = 0; i < 2; i++) {
            var v = sub(hs[i], c[0]);
            if (Math.abs(cross(u, v)) > TOL * 0.5) { return false; }
            var al = dot(u, v);
            if (al < -TOL || al > d + TOL) { return false; }
        }
        return true;
    }
    function startTan(c) {
        var v = sub(c[1], c[0]);
        if (vlen(v) < 1e-9) { v = sub(c[2], c[0]); }
        if (vlen(v) < 1e-9) { v = sub(c[3], c[0]); }
        return norm(v);
    }
    function endTan(c) {
        var v = sub(c[3], c[2]);
        if (vlen(v) < 1e-9) { v = sub(c[3], c[1]); }
        if (vlen(v) < 1e-9) { v = sub(c[3], c[0]); }
        return norm(v);
    }
    function reverseCubic(c) { return [cp(c[3]), cp(c[2]), cp(c[1]), cp(c[0])]; }
    function reverseChain(ch) {
        var r = [];
        for (var i = ch.length - 1; i >= 0; i--) { r.push(reverseCubic(ch[i])); }
        return r;
    }
    function lineCubic(a, b) { return [cp(a), cp(a), cp(b), cp(b)]; }

    function pointAtU(ch, u) {
        var i = Math.floor(u), t = u - i;
        if (i >= ch.length) { i = ch.length - 1; t = 1; }
        if (i < 0) { i = 0; t = 0; }
        return bez(ch[i], t);
    }
    function subChain(ch, u0, u1) {
        var i0 = Math.floor(u0), t0 = u0 - i0;
        var i1 = Math.floor(u1), t1 = u1 - i1;
        if (i0 >= ch.length) { i0 = ch.length - 1; t0 = 1; }
        if (t1 < 1e-9 && i1 > 0) { i1--; t1 = 1; }
        if (i1 >= ch.length) { i1 = ch.length - 1; t1 = 1; }
        var r = [];
        for (var i = i0; i <= i1; i++) {
            var a = (i === i0) ? t0 : 0;
            var b = (i === i1) ? t1 : 1;
            if (b - a > 1e-9) { r.push(subCubic(ch[i], a, b)); }
        }
        return r;
    }

    function ctrlLen(c) { return dist(c[0], c[1]) + dist(c[1], c[2]) + dist(c[2], c[3]); }
    function bboxOf(pts) {
        var b = [pts[0][0], pts[0][1], pts[0][0], pts[0][1]];
        for (var i = 1; i < pts.length; i++) {
            if (pts[i][0] < b[0]) { b[0] = pts[i][0]; }
            if (pts[i][1] < b[1]) { b[1] = pts[i][1]; }
            if (pts[i][0] > b[2]) { b[2] = pts[i][0]; }
            if (pts[i][1] > b[3]) { b[3] = pts[i][1]; }
        }
        return b;
    }
    function bbOverlap(a, b, pad) {
        return !(a[2] + pad < b[0] || b[2] + pad < a[0] || a[3] + pad < b[1] || b[3] + pad < a[1]);
    }
    function fracToT(c, f) {
        // Straight cubic: bezier t at which the point is fraction f of the way along
        if (f <= 0) { return 0; }
        if (f >= 1) { return 1; }
        var d = sub(c[3], c[0]), L2 = dot(d, d);
        if (L2 < 1e-18) { return f; }
        var lo = 0, hi = 1;
        for (var it = 0; it < 40; it++) {
            var mid = (lo + hi) / 2;
            if (dot(sub(bez(c, mid), c[0]), d) / L2 < f) { lo = mid; } else { hi = mid; }
        }
        return (lo + hi) / 2;
    }
    function sampleChain(ch) {
        var pts = [], us = [], lin = [];
        for (var i = 0; i < ch.length; i++) {
            var c = ch[i];
            if (i === 0) { pts.push(cp(c[0])); us.push(0); }
            if (isLinear(c)) {
                pts.push(cp(c[3])); us.push(i + 1); lin.push(i);
            } else {
                var n = Math.ceil(ctrlLen(c) / (0.05 * IN));
                n = Math.max(12, Math.min(64, n));
                for (var k = 1; k <= n; k++) { pts.push(bez(c, k / n)); us.push(i + k / n); lin.push(-1); }
            }
        }
        return { pts: pts, us: us, lin: lin, ch: ch, bb: bboxOf(pts) };
    }
    function polyLen(pts) {
        var l = 0;
        for (var i = 1; i < pts.length; i++) { l += dist(pts[i - 1], pts[i]); }
        return l;
    }
    function polyArea(pts) {
        var a = 0;
        for (var i = 0; i < pts.length; i++) {
            var p = pts[i], q = pts[(i + 1) % pts.length];
            a += p[0] * q[1] - q[0] * p[1];
        }
        return a / 2;
    }
    function pointInPoly(pt, pts, bb) {
        if (bb && (pt[0] < bb[0] || pt[0] > bb[2] || pt[1] < bb[1] || pt[1] > bb[3])) { return false; }
        var inside = false;
        for (var i = 0, j = pts.length - 1; i < pts.length; j = i++) {
            var xi = pts[i][0], yi = pts[i][1], xj = pts[j][0], yj = pts[j][1];
            if (((yi > pt[1]) !== (yj > pt[1])) &&
                (pt[0] < (xj - xi) * (pt[1] - yi) / (yj - yi) + xi)) {
                inside = !inside;
            }
        }
        return inside;
    }
    function insidePoint(pts, area) {
        for (var i = 0; i < pts.length - 1; i++) {
            var a = pts[i], b = pts[i + 1];
            var d = sub(b, a), l = vlen(d);
            if (l < TOL) { continue; }
            var n = [-d[1] / l, d[0] / l];
            if (area < 0) { n = [-n[0], -n[1]]; }
            return add(lerp(a, b, 0.5), mul(n, Math.min(TOL, l * 0.25)));
        }
        return pts[0];
    }
    function segInter(p, p2, q, q2) {
        var r = sub(p2, p), s = sub(q2, q);
        var den = cross(r, s);
        if (Math.abs(den) < 1e-12) { return null; }
        var qp = sub(q, p);
        return [cross(qp, s) / den, cross(qp, r) / den];
    }
    function pointSegDist(pt, a, b) {
        var d = sub(b, a), l2 = dot(d, d);
        var t = l2 < 1e-18 ? 0 : clamp(dot(sub(pt, a), d) / l2, 0, 1);
        return { d: dist(pt, lerp(a, b, t)), t: t };
    }
    function isStraightChain(ch) {
        var a = ch[0][0], b = ch[ch.length - 1][3];
        if (dist(a, b) < TOL) { return false; }
        var u = norm(sub(b, a));
        for (var i = 0; i < ch.length; i++) {
            if (!isLinear(ch[i])) { return false; }
            if (Math.abs(cross(u, sub(ch[i][3], a))) > TOL * 0.5) { return false; }
            if (dot(u, sub(ch[i][3], ch[i][0])) < -TOL) { return false; }
        }
        return true;
    }

    // ------------------------------------------------------------------
    // Illustrator helpers
    // ------------------------------------------------------------------
    function chainFromPath(p) {
        var pts = p.pathPoints, n = pts.length;
        var A = [], L = [], R = [];
        for (var i = 0; i < n; i++) {
            A.push(cp(pts[i].anchor));
            L.push(cp(pts[i].leftDirection));
            R.push(cp(pts[i].rightDirection));
        }
        var segCount = p.closed ? n : n - 1;
        var ch = [];
        for (var s = 0; s < segCount; s++) {
            var j = (s + 1) % n;
            var c = [cp(A[s]), cp(R[s]), cp(L[j]), cp(A[j])];
            if (dist(c[0], c[3]) < 1e-6 && dist(c[0], c[1]) < 1e-6 && dist(c[2], c[3]) < 1e-6) { continue; }
            ch.push(c);
        }
        return ch;
    }
    function spotNameOf(p) {
        try { if (p.stroked && p.strokeColor.typename === "SpotColor") { return p.strokeColor.spot.name; } } catch (e1) {}
        try { if (p.filled && p.fillColor.typename === "SpotColor") { return p.fillColor.spot.name; } } catch (e2) {}
        return null;
    }
    function getSpot(name, cmyk) {
        for (var i = 0; i < doc.spots.length; i++) {
            if (doc.spots[i].name === name) { return doc.spots[i]; }
        }
        var s = doc.spots.add();
        s.name = name;
        var c = new CMYKColor();
        c.cyan = cmyk[0]; c.magenta = cmyk[1]; c.yellow = cmyk[2]; c.black = cmyk[3];
        s.color = c;
        s.colorType = ColorModel.SPOT;
        return s;
    }
    function spotColorOf(spot) {
        var c = new SpotColor();
        c.spot = spot;
        c.tint = 100;
        return c;
    }
    function findLayer(name) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === name) { return doc.layers[i]; }
        }
        return null;
    }
    function getOrCreateLayer(name) {
        var lyr = findLayer(name);
        if (lyr) { return lyr; }
        var regLayer = findLayer("REG");
        var nl = doc.layers.add();
        nl.name = name;
        if (regLayer) { nl.move(regLayer, ElementPlacement.PLACEAFTER); }
        return nl;
    }
    var layerStates = [];
    function openLayer(lyr) {
        for (var i = 0; i < layerStates.length; i++) {
            if (layerStates[i].layer === lyr) { return; }
        }
        layerStates.push({ layer: lyr, locked: lyr.locked, visible: lyr.visible });
        if (!lyr.visible) { lyr.visible = true; }
        if (lyr.locked) { lyr.locked = false; }
    }
    function restoreLayers() {
        for (var i = 0; i < layerStates.length; i++) {
            try {
                layerStates[i].layer.visible = layerStates[i].visible;
                layerStates[i].layer.locked = layerStates[i].locked;
            } catch (e) {}
        }
    }
    function isSmoothJoin(l, a, r) {
        var v1 = sub(a, l), v2 = sub(r, a);
        if (vlen(v1) < 1e-6 || vlen(v2) < 1e-6) { return false; }
        var n1 = norm(v1), n2 = norm(v2);
        return dot(n1, n2) > 0.9999;
    }
    function makePath(layer, ch, closed, color, strokeW) {
        var anchors = [], L = [], R = [], n = ch.length, i;
        if (closed) {
            for (i = 0; i < n; i++) {
                anchors.push(cp(ch[i][0]));
                L.push(cp(ch[(i - 1 + n) % n][2]));
                R.push(cp(ch[i][1]));
            }
        } else {
            anchors.push(cp(ch[0][0])); L.push(cp(ch[0][0])); R.push(cp(ch[0][1]));
            for (i = 0; i < n; i++) {
                anchors.push(cp(ch[i][3]));
                L.push(cp(ch[i][2]));
                R.push(i + 1 < n ? cp(ch[i + 1][1]) : cp(ch[i][3]));
            }
        }
        var p = layer.pathItems.add();
        p.setEntirePath(anchors);
        for (i = 0; i < anchors.length; i++) {
            var pp = p.pathPoints[i];
            pp.pointType = isSmoothJoin(L[i], anchors[i], R[i]) ? PointType.SMOOTH : PointType.CORNER;
            pp.leftDirection = L[i];
            pp.rightDirection = R[i];
        }
        p.closed = closed;
        p.filled = false;
        p.stroked = true;
        p.strokeColor = color;
        p.strokeWidth = BKC_STROKE_PT / sf;
        return p;
    }

    // ------------------------------------------------------------------
    // Gather selection
    // ------------------------------------------------------------------
    var bitItems = [], knifeItems = [], ignored = 0, cidCounter = 0;
    function pushPath(p, compound, cid) {
        var name = spotNameOf(p);
        if (!name && compound) {
            try { if (compound.pathItems.length > 0) { name = spotNameOf(compound.pathItems[0]); } } catch (e0) {}
        }
        var sw = 0;
        try { sw = p.strokeWidth; } catch (e1) {}
        var rec = { path: p, src: compound ? compound : p, cid: cid, color: name, strokeW: sw };
        if (name === BKC_COLOR_OUTSIDE || name === BKC_COLOR_INSIDE) { bitItems.push(rec); }
        else if (name === BKC_COLOR_KNIFE) { knifeItems.push(rec); }
        else { ignored++; }
    }
    function collect(item) {
        var t = item.typename;
        if (t === "PathItem") {
            cidCounter++;
            pushPath(item, null, cidCounter);
        } else if (t === "GroupItem") {
            for (var i = 0; i < item.pageItems.length; i++) { collect(item.pageItems[i]); }
        } else if (t === "CompoundPathItem") {
            cidCounter++;
            for (var j = 0; j < item.pathItems.length; j++) { pushPath(item.pathItems[j], item, cidCounter); }
        } else {
            ignored++;
        }
    }
    for (var si = 0; si < sel.length; si++) { collect(sel[si]); }

    if (bitItems.length > 0 && knifeItems.length > 0) {
        alert("Selection has both bit cut (" + BKC_COLOR_OUTSIDE + " / " + BKC_COLOR_INSIDE + ") and knife cut (" +
              BKC_COLOR_KNIFE + ") paths.\nSelect only one kind and run again.");
        return;
    }
    if (bitItems.length === 0 && knifeItems.length === 0) {
        alert("No cut paths found in the selection.\nLooking for strokes in: " +
              BKC_COLOR_OUTSIDE + ", " + BKC_COLOR_INSIDE + ", " + BKC_COLOR_KNIFE);
        return;
    }

    function deleteSources(recs) {
        var gone = {};
        for (var i = 0; i < recs.length; i++) {
            if (!recs[i].used) { continue; }
            if (gone[recs[i].cid]) { continue; }
            gone[recs[i].cid] = true;
            try { recs[i].src.remove(); } catch (e) {}
        }
    }

    var newItems = [];
    try {
        if (bitItems.length > 0) {
            bitToKnife();
        } else {
            knifeToBit();
        }
    } finally {
        restoreLayers();
    }
    try { doc.selection = newItems; } catch (eSel) {}

    // ==================================================================
    // BIT -> KNIFE
    // ==================================================================
    function bitToKnife() {
        var recs = [], skippedOpen = 0, i, j;
        for (i = 0; i < bitItems.length; i++) {
            var b = bitItems[i];
            var ch = chainFromPath(b.path);
            if (!b.path.closed || ch.length < 1) { skippedOpen++; continue; }
            var s = sampleChain(ch);
            b.ch = ch;
            b.pts = s.pts;
            b.bb = s.bb;
            b.area = polyArea(s.pts);
            if (Math.abs(b.area) < TOL * TOL) { skippedOpen++; continue; }
            recs.push(b);
        }
        if (recs.length === 0) {
            alert("No closed bit cut paths found in the selection.");
            return;
        }

        // Hole detection: Inside color, flipped by nesting inside the same compound path
        for (i = 0; i < recs.length; i++) {
            var depth = 0;
            var probe = insidePoint(recs[i].pts, recs[i].area);
            for (j = 0; j < recs.length; j++) {
                if (i === j || recs[j].cid !== recs[i].cid) { continue; }
                if (pointInPoly(probe, recs[j].pts, recs[j].bb)) { depth++; }
            }
            recs[i].isHole = (recs[i].color === BKC_COLOR_INSIDE) !== (depth % 2 === 1);
        }

        // Obstacles: selected geometry plus everything on the cut layers
        var obstacles = [];
        for (i = 0; i < recs.length; i++) { obstacles.push({ pts: recs[i].pts, bb: recs[i].bb }); }
        for (i = 0; i < BKC_OBSTACLE_LAYERS.length; i++) {
            var lyr = findLayer(BKC_OBSTACLE_LAYERS[i]);
            if (!lyr) { continue; }
            for (j = 0; j < lyr.pathItems.length; j++) {
                try {
                    var op = lyr.pathItems[j];
                    if (op.guides) { continue; }
                    var och = chainFromPath(op);
                    if (och.length === 0) { continue; }
                    var os = sampleChain(och);
                    obstacles.push({ pts: os.pts, bb: os.bb });
                } catch (eObs) {}
            }
        }

        function inBody(pt) {
            var solid = 0, holes = 0;
            for (var k = 0; k < recs.length; k++) {
                if (pointInPoly(pt, recs[k].pts, recs[k].bb)) {
                    if (recs[k].isHole) { holes++; } else { solid++; }
                }
            }
            return solid > holes;
        }
        function firstHit(E, F, skipStart, skipEnd) {
            // Distance from E along E->F to the first obstacle crossing, or -1
            var L = dist(E, F);
            if (L < 1e-9) { return -1; }
            var bb = [Math.min(E[0], F[0]), Math.min(E[1], F[1]), Math.max(E[0], F[0]), Math.max(E[1], F[1])];
            var best = -1;
            for (var k = 0; k < obstacles.length; k++) {
                var o = obstacles[k];
                if (!bbOverlap(bb, o.bb, TOL)) { continue; }
                for (var m = 0; m < o.pts.length - 1; m++) {
                    var a = o.pts[m], c = o.pts[m + 1];
                    if (Math.max(a[0], c[0]) < bb[0] - TOL || Math.min(a[0], c[0]) > bb[2] + TOL ||
                        Math.max(a[1], c[1]) < bb[1] - TOL || Math.min(a[1], c[1]) > bb[3] + TOL) { continue; }
                    var r = segInter(E, F, a, c);
                    if (!r) { continue; }
                    if (r[1] < -1e-9 || r[1] > 1 + 1e-9 || r[0] < 0 || r[0] > 1) { continue; }
                    var sd = r[0] * L;
                    if (sd < skipStart || sd > L - skipEnd) { continue; }
                    if (best < 0 || sd < best) { best = sd; }
                }
            }
            return best;
        }
        var stats = { shortened: 0, dropped: 0 };
        function extLen(E, d, flagged) {
            if (!flagged) { return 0; }
            if (inBody(add(E, mul(d, PROBE)))) { stats.dropped++; return 0; }
            var hit = firstHit(E, add(E, mul(d, OV)), TOL * 1.5, 0);
            if (hit < 0) { return OV; }
            var l = hit - MARGIN;
            if (l < MIN_EXT) { stats.dropped++; return 0; }
            stats.shortened++;
            return l;
        }

        // Split each path at sharp corners
        var curves = [], lines = [], loops = [];
        for (i = 0; i < recs.length; i++) {
            var r = recs[i];
            var c = r.ch, m = c.length;
            var corners = [];
            for (j = 0; j < m; j++) {
                var tin = endTan(c[(j - 1 + m) % m]);
                var tout = startTan(c[j]);
                var ang = Math.acos(clamp(dot(tin, tout), -1, 1));
                if (ang > CORNER_RAD) {
                    var V = c[j][0];
                    var pIn = add(V, mul(tin, PROBE));
                    var pOut = sub(V, mul(tout, PROBE));
                    var inIn = pointInPoly(pIn, r.pts, r.bb);
                    var inOut = pointInPoly(pOut, r.pts, r.bb);
                    corners.push({ i: j, extIn: r.isHole ? inIn : !inIn, extOut: r.isHole ? inOut : !inOut });
                }
            }
            r.used = true;
            if (corners.length === 0) {
                loops.push({ ch: c, rec: r });
                continue;
            }
            for (var k = 0; k < corners.length; k++) {
                var c0 = corners[k], c1 = corners[(k + 1) % corners.length];
                var segs = [];
                var idx = c0.i;
                do {
                    segs.push(c[idx]);
                    idx = (idx + 1) % m;
                } while (idx !== c1.i);
                var piece = { ch: segs, startExt: c0.extOut, endExt: c1.extIn, rec: r };
                if (isStraightChain(segs)) {
                    lines.push({ A: segs[0][0], B: segs[segs.length - 1][3], aExt: piece.startExt, bExt: piece.endExt, rec: r });
                } else {
                    curves.push(piece);
                }
            }
        }

        // Merge colinear straight edges whose overshoots touch
        var clusters = [];
        for (i = 0; i < lines.length; i++) {
            var ln = lines[i];
            var u = norm(sub(ln.B, ln.A));
            var cl = null;
            for (j = 0; j < clusters.length; j++) {
                var q = clusters[j];
                if (q.hole !== ln.rec.isHole) { continue; }
                if (Math.abs(cross(u, q.u)) > ANG_EPS) { continue; }
                if (Math.abs(dot(sub(ln.A, q.P), q.n)) > TOL || Math.abs(dot(sub(ln.B, q.P), q.n)) > TOL) { continue; }
                cl = q;
                break;
            }
            if (!cl) {
                cl = { P: cp(ln.A), u: u, n: [-u[1], u[0]], hole: ln.rec.isHole, items: [] };
                clusters.push(cl);
            }
            var ta = dot(sub(ln.A, cl.P), cl.u), tb = dot(sub(ln.B, cl.P), cl.u);
            if (ta <= tb) { cl.items.push({ lo: ta, hi: tb, loExt: ln.aExt, hiExt: ln.bExt, rec: ln.rec }); }
            else { cl.items.push({ lo: tb, hi: ta, loExt: ln.bExt, hiExt: ln.aExt, rec: ln.rec }); }
        }
        var mergedLines = [], mergeCount = 0;
        function reachLo(it) { return it.lo - (it.loExt ? OV : 0); }
        function reachHi(it) { return it.hi + (it.hiExt ? OV : 0); }
        for (i = 0; i < clusters.length; i++) {
            var cq = clusters[i];
            cq.items.sort(function (a, b2) { return reachLo(a) - reachLo(b2); });
            var cur = null;
            for (j = 0; j < cq.items.length; j++) {
                var it = cq.items[j];
                if (!cur) { cur = { lo: it.lo, hi: it.hi, loExt: it.loExt, hiExt: it.hiExt, rec: it.rec, cl: cq }; continue; }
                var ok = reachLo(it) <= reachHi(cur) + TOL;
                if (ok && it.lo > cur.hi + TOL) {
                    var g0 = add(cq.P, mul(cq.u, cur.hi)), g1 = add(cq.P, mul(cq.u, it.lo));
                    if (inBody(lerp(g0, g1, 0.5)) || firstHit(g0, g1, TOL * 1.5, TOL * 1.5) >= 0) { ok = false; }
                }
                if (ok) {
                    mergeCount++;
                    if (it.lo < cur.lo - TOL) { cur.lo = it.lo; cur.loExt = it.loExt; }
                    else if (Math.abs(it.lo - cur.lo) <= TOL) { cur.loExt = cur.loExt && it.loExt; }
                    if (it.hi > cur.hi + TOL) { cur.hi = it.hi; cur.hiExt = it.hiExt; }
                    else if (Math.abs(it.hi - cur.hi) <= TOL) { cur.hiExt = cur.hiExt && it.hiExt; }
                } else {
                    mergedLines.push(cur);
                    cur = { lo: it.lo, hi: it.hi, loExt: it.loExt, hiExt: it.hiExt, rec: it.rec, cl: cq };
                }
            }
            if (cur) { mergedLines.push(cur); }
        }

        // Build output chains with overshoots
        var outputs = []; // {ch, closed, hole, strokeW}
        for (i = 0; i < mergedLines.length; i++) {
            var ml = mergedLines[i];
            var A = add(ml.cl.P, mul(ml.cl.u, ml.lo));
            var B = add(ml.cl.P, mul(ml.cl.u, ml.hi));
            var la = extLen(A, mul(ml.cl.u, -1), ml.loExt);
            var lb = extLen(B, ml.cl.u, ml.hiExt);
            var A2 = sub(A, mul(ml.cl.u, la)), B2 = add(B, mul(ml.cl.u, lb));
            outputs.push({ ch: [lineCubic(A2, B2)], closed: false, hole: ml.cl.hole, strokeW: ml.rec.strokeW });
        }
        for (i = 0; i < curves.length; i++) {
            var pc = curves[i];
            var out = pc.ch.slice(0);
            var S = out[0][0], dS = mul(startTan(out[0]), -1);
            var E = out[out.length - 1][3], dE = endTan(out[out.length - 1]);
            var ls = extLen(S, dS, pc.startExt);
            var le = extLen(E, dE, pc.endExt);
            if (ls > 0) { out.unshift(lineCubic(add(S, mul(dS, ls)), S)); }
            if (le > 0) { out.push(lineCubic(E, add(E, mul(dE, le)))); }
            outputs.push({ ch: out, closed: false, hole: pc.rec.isHole, strokeW: pc.rec.strokeW });
        }
        for (i = 0; i < loops.length; i++) {
            outputs.push({ ch: loops[i].ch, closed: true, hole: loops[i].rec.isHole, strokeW: loops[i].rec.strokeW });
        }

        // Stacking: pathItems.add() puts new items on top, so create the bottom ones first
        var firstHole = BKC_INNER_CUTS_AT_BOTTOM;
        var layer = getOrCreateLayer(BKC_COLOR_KNIFE);
        openLayer(layer);
        var color = spotColorOf(getSpot(BKC_COLOR_KNIFE, [100, 0, 100, 0]));
        for (var pass = 0; pass < 2; pass++) {
            var wantHole = (pass === 0) ? firstHole : !firstHole;
            for (i = 0; i < outputs.length; i++) {
                if (outputs[i].hole !== wantHole) { continue; }
                newItems.push(makePath(layer, outputs[i].ch, outputs[i].closed, color, outputs[i].strokeW));
            }
        }
        deleteSources(recs);

        var msg = "Bit cut -> knife cut\n\n" +
                  recs.length + " path(s) converted into " + newItems.length + " knife cut(s)\n" +
                  mergeCount + " colinear edge merge(s)\n";
        if (stats.shortened > 0) { msg += stats.shortened + " overshoot(s) shortened to clear other cuts\n"; }
        if (stats.dropped > 0) { msg += stats.dropped + " overshoot(s) removed (no room)\n"; }
        if (skippedOpen > 0) { msg += skippedOpen + " open/empty bit path(s) skipped (left as-is)\n"; }
        if (ignored > 0) { msg += ignored + " non-cut item(s) ignored\n"; }
        alert(msg);
    }

    // ==================================================================
    // KNIFE -> BIT
    // ==================================================================
    function knifeToBit() {
        var opens = [], faces = [], i, j, k;
        for (i = 0; i < knifeItems.length; i++) {
            var kr = knifeItems[i];
            var ch = chainFromPath(kr.path);
            if (ch.length === 0) { continue; }
            if (kr.path.closed) {
                kr.used = true;
                faces.push({ ch: ch, strokeW: kr.strokeW });
            } else {
                kr.ch = ch;
                opens.push(kr);
            }
        }

        function computeHits(list) {
            var hits = [];
            for (var a = 0; a < list.length; a++) { hits.push([]); }
            for (a = 0; a < list.length; a++) {
                for (var b = a; b < list.length; b++) {
                    var A = list[a].s, B = list[b].s, same = (a === b);
                    if (!same && !bbOverlap(A.bb, B.bb, TOL)) { continue; }
                    var pa = A.pts, pb = B.pts;
                    for (var x = 0; x < pa.length - 1; x++) {
                        var a0 = pa[x], a1 = pa[x + 1];
                        var la = dist(a0, a1);
                        if (la < 1e-9) { continue; }
                        var bxMin = Math.min(a0[0], a1[0]) - TOL, bxMax = Math.max(a0[0], a1[0]) + TOL;
                        var byMin = Math.min(a0[1], a1[1]) - TOL, byMax = Math.max(a0[1], a1[1]) + TOL;
                        for (var y = same ? x + 2 : 0; y < pb.length - 1; y++) {
                            var b0 = pb[y], b1 = pb[y + 1];
                            if (Math.max(b0[0], b1[0]) < bxMin || Math.min(b0[0], b1[0]) > bxMax ||
                                Math.max(b0[1], b1[1]) < byMin || Math.min(b0[1], b1[1]) > byMax) { continue; }
                            var lb = dist(b0, b1);
                            if (lb < 1e-9) { continue; }
                            var r = segInter(a0, a1, b0, b1);
                            if (!r) { continue; }
                            var da = TOL / la, db = TOL / lb;
                            if (r[0] < -da || r[0] > 1 + da || r[1] < -db || r[1] > 1 + db) { continue; }
                            var t0 = clamp(r[0], 0, 1), t1 = clamp(r[1], 0, 1);
                            var p = lerp(a0, a1, t0);
                            hits[a].push({ u: uOf(A, x, t0), p: p });
                            hits[b].push({ u: uOf(B, y, t1), p: p });
                        }
                    }
                    if (!same) {
                        touches(list[a].s, list[b].s, hits[a], hits[b]);
                        touches(list[b].s, list[a].s, hits[b], hits[a]);
                    }
                }
            }
            return hits;
        }
        function uOf(S, x, t) {
            if (S.lin[x] >= 0) { return S.lin[x] + fracToT(S.ch[S.lin[x]], t); }
            return S.us[x] + (S.us[x + 1] - S.us[x]) * t;
        }
        function touches(P, Q, hp, hq) {
            var ends = [0, P.pts.length - 1];
            for (var e = 0; e < 2; e++) {
                var ep = P.pts[ends[e]];
                if (ep[0] < Q.bb[0] - TOL || ep[0] > Q.bb[2] + TOL || ep[1] < Q.bb[1] - TOL || ep[1] > Q.bb[3] + TOL) { continue; }
                for (var y = 0; y < Q.pts.length - 1; y++) {
                    var r = pointSegDist(ep, Q.pts[y], Q.pts[y + 1]);
                    if (r.d < TOL) {
                        hp.push({ u: P.us[ends[e]], p: cp(ep) });
                        hq.push({ u: uOf(Q, y, r.t), p: cp(ep) });
                        break;
                    }
                }
            }
        }

        // 1. Trim overshoot stubs off free (untouching) straight ends
        for (i = 0; i < opens.length; i++) { opens[i].s = sampleChain(opens[i].ch); }
        var hits = computeHits(opens);
        var trimmed = 0;
        var WIN = OV * BKC_TRIM_WINDOW_FACTOR + TOL;
        for (i = 0; i < opens.length; i++) {
            var o = opens[i], c = o.ch, m = c.length, h = hits[i];
            var S = c[0][0], E = c[m - 1][3];
            var u0 = 0, u1 = m;
            var startFree = true, endFree = true;
            for (k = 0; k < h.length; k++) {
                if (dist(h[k].p, S) <= TOL * 1.5) { startFree = false; }
                if (dist(h[k].p, E) <= TOL * 1.5) { endFree = false; }
            }
            if (startFree && isLinear(c[0])) { u0 = pickTrim(h, S, 0, 1, true); }
            if (endFree && isLinear(c[m - 1])) { u1 = pickTrim(h, E, m - 1, m, false); }
            if (u0 > 0) { trimmed++; }
            if (u1 < m) { trimmed++; }
            if (u1 - u0 > 1e-9 && (u0 > 0 || u1 < m)) {
                o.ch = subChain(c, u0, u1);
            }
        }
        function pickTrim(h, tip, uMin, uMax, atStart) {
            // Farthest hit within OV of the tip; otherwise nearest within the window
            var best = -1, bestD = -1, near = -1, nearD = 1e18;
            for (var q = 0; q < h.length; q++) {
                if (h[q].u < uMin || h[q].u > uMax) { continue; }
                var d = dist(h[q].p, tip);
                if (d <= TOL * 1.5 || d > WIN) { continue; }
                if (d <= OV + TOL && d > bestD) { bestD = d; best = h[q].u; }
                if (d < nearD) { nearD = d; near = h[q].u; }
            }
            var pick = best >= 0 ? best : near;
            if (pick < 0) { return atStart ? 0 : uMax; }
            return pick;
        }

        // 2. Split at every crossing / touch and build a planar graph
        for (i = 0; i < opens.length; i++) { opens[i].s = sampleChain(opens[i].ch); }
        hits = computeHits(opens);
        var nodes = [], edges = [];
        function nodeAt(p) {
            for (var q = 0; q < nodes.length; q++) {
                if (dist(nodes[q].p, p) <= TOL * 1.5) { return q; }
            }
            nodes.push({ p: cp(p), out: [] });
            return nodes.length - 1;
        }
        for (i = 0; i < opens.length; i++) {
            var oc = opens[i].ch, om = oc.length, hh = hits[i];
            hh.sort(function (a, b) { return a.u - b.u; });
            var startP = oc[0][0], endP = oc[om - 1][3];
            var splits = [0], lastP = startP;
            for (k = 0; k < hh.length; k++) {
                var hp = pointAtU(oc, hh[k].u);
                if (dist(hp, lastP) <= TOL * 1.5 || dist(hp, endP) <= TOL * 1.5) { continue; }
                if (hh[k].u <= splits[splits.length - 1] || hh[k].u >= om) { continue; }
                splits.push(hh[k].u);
                lastP = hp;
            }
            splits.push(om);
            for (k = 0; k < splits.length - 1; k++) {
                var sc = subChain(oc, splits[k], splits[k + 1]);
                if (sc.length === 0) { continue; }
                var ssm = sampleChain(sc);
                var elen = polyLen(ssm.pts);
                if (elen < TOL) { continue; }
                var na = nodeAt(sc[0][0]), nb = nodeAt(sc[sc.length - 1][3]);
                // Snap ends to node positions
                var d0 = sub(nodes[na].p, sc[0][0]);
                sc[0][0] = add(sc[0][0], d0); sc[0][1] = add(sc[0][1], d0);
                var last = sc[sc.length - 1];
                var d1 = sub(nodes[nb].p, last[3]);
                last[3] = add(last[3], d1); last[2] = add(last[2], d1);
                var interiorShort = (k > 0 && k < splits.length - 2 && isStraightChain(sc) && elen <= 2 * OV + TOL * 2);
                edges.push({ a: na, b: nb, ch: sc, len: elen, alive: true, interiorShort: interiorShort, src: opens[i] });
            }
        }

        // 3. Prune dangling edges (leftover stubs / open cuts)
        var deg = [];
        for (i = 0; i < nodes.length; i++) { deg.push(0); }
        for (i = 0; i < edges.length; i++) { deg[edges[i].a]++; deg[edges[i].b]++; }
        var longDangling = 0, changed = true;
        while (changed) {
            changed = false;
            for (i = 0; i < edges.length; i++) {
                var ed = edges[i];
                if (!ed.alive) { continue; }
                if (deg[ed.a] === 1 || deg[ed.b] === 1) {
                    ed.alive = false;
                    deg[ed.a]--; deg[ed.b]--;
                    if (ed.len > WIN) { longDangling++; }
                    changed = true;
                }
            }
        }

        // 4. Faces: half-edges sorted by angle around each node
        var halves = [];
        function dirAngle(chn, atStart) {
            var p0, p1;
            if (atStart) { p0 = chn[0][0]; p1 = bez(chn[0], 0.02); }
            else { var lc = chn[chn.length - 1]; p0 = lc[3]; p1 = bez(lc, 0.98); }
            return Math.atan2(p1[1] - p0[1], p1[0] - p0[0]);
        }
        for (i = 0; i < edges.length; i++) {
            var e2 = edges[i];
            if (!e2.alive) { continue; }
            var h0 = { e: e2, fwd: true, from: e2.a, to: e2.b, ang: dirAngle(e2.ch, true), seen: false };
            var h1 = { e: e2, fwd: false, from: e2.b, to: e2.a, ang: dirAngle(e2.ch, false), seen: false };
            h0.twin = h1; h1.twin = h0;
            halves.push(h0); halves.push(h1);
            nodes[h0.from].out.push(h0);
            nodes[h1.from].out.push(h1);
        }
        for (i = 0; i < nodes.length; i++) {
            nodes[i].out.sort(function (a, b) { return a.ang - b.ang; });
            for (k = 0; k < nodes[i].out.length; k++) { nodes[i].out[k].idx = k; }
        }
        var gapSkipped = 0;
        var minArea = (0.05 * IN) * (0.05 * IN);
        for (i = 0; i < halves.length; i++) {
            if (halves[i].seen) { continue; }
            var cyc = [], hcur = halves[i], guard = 0;
            while (!hcur.seen && guard < 100000) {
                hcur.seen = true;
                cyc.push(hcur);
                var lst = nodes[hcur.to].out;
                hcur = lst[(hcur.twin.idx - 1 + lst.length) % lst.length];
                guard++;
            }
            var fch = [], hasShort = false;
            for (k = 0; k < cyc.length; k++) {
                var part = cyc[k].fwd ? cyc[k].e.ch : reverseChain(cyc[k].e.ch);
                for (j = 0; j < part.length; j++) { fch.push(part[j]); }
                if (cyc[k].e.interiorShort) { hasShort = true; }
            }
            if (fch.length === 0) { continue; }
            var fs = sampleChain(fch);
            var area = polyArea(fs.pts);
            if (area <= minArea) { continue; }
            var bw = fs.bb[2] - fs.bb[0], bh = fs.bb[3] - fs.bb[1];
            if (hasShort && Math.min(bw, bh) <= 2 * OV + TOL * 2) { gapSkipped++; continue; }
            for (k = 0; k < cyc.length; k++) { cyc[k].e.src.used = true; }
            faces.push({ ch: fch, strokeW: cyc[0].e.src.strokeW });
        }

        if (faces.length === 0) {
            alert("Could not rebuild any closed shapes from the selected knife cuts.\nNothing was changed.");
            return;
        }

        // 5. Inside / Outside by nesting
        for (i = 0; i < faces.length; i++) {
            var fsi = sampleChain(faces[i].ch);
            faces[i].pts = fsi.pts;
            faces[i].bb = fsi.bb;
            faces[i].area = polyArea(fsi.pts);
        }
        for (i = 0; i < faces.length; i++) {
            var probe = insidePoint(faces[i].pts, faces[i].area);
            var depth = 0;
            for (j = 0; j < faces.length; j++) {
                if (i === j) { continue; }
                if (pointInPoly(probe, faces[j].pts, faces[j].bb)) { depth++; }
            }
            faces[i].inside = (depth % 2 === 1);
        }

        // 6. Output
        var outLayer = getOrCreateLayer(BKC_COLOR_OUTSIDE);
        var inLayer = null;
        openLayer(outLayer);
        var outColor = spotColorOf(getSpot(BKC_COLOR_OUTSIDE, [0, 100, 0, 0]));
        var inColor = null;
        for (i = 0; i < faces.length; i++) {
            var fc = simplifyClosed(faces[i].ch);
            if (faces[i].inside) {
                if (!inLayer) {
                    inLayer = getOrCreateLayer(BKC_COLOR_INSIDE);
                    openLayer(inLayer);
                    inColor = spotColorOf(getSpot(BKC_COLOR_INSIDE, [100, 0, 0, 0]));
                }
                newItems.push(makePath(inLayer, fc, true, inColor, faces[i].strokeW));
            } else {
                newItems.push(makePath(outLayer, fc, true, outColor, faces[i].strokeW));
            }
        }
        deleteSources(knifeItems);

        var unused = 0;
        for (i = 0; i < knifeItems.length; i++) { if (!knifeItems[i].used) { unused++; } }
        var msg = "Knife cut -> bit cut\n\n" +
                  newItems.length + " closed shape(s) rebuilt\n" +
                  trimmed + " overshoot(s) trimmed\n";
        if (gapSkipped > 0) { msg += gapSkipped + " narrow gap region(s) between signs skipped\n"; }
        if (longDangling > 0) { msg += longDangling + " loose cut segment(s) longer than an overshoot were dropped - check result\n"; }
        if (unused > 0) { msg += unused + " knife path(s) did not form a closed shape and were left in place\n"; }
        if (ignored > 0) { msg += ignored + " non-cut item(s) ignored\n"; }
        alert(msg);
    }

    function simplifyClosed(ch) {
        // Merge consecutive straight, colinear segments
        var r = [];
        for (var i = 0; i < ch.length; i++) {
            var c = ch[i];
            if (r.length > 0) {
                var p = r[r.length - 1];
                if (isLinear(p) && isLinear(c)) {
                    var d1 = norm(sub(p[3], p[0])), d2 = norm(sub(c[3], c[0]));
                    if (dot(d1, d2) > 0.99999) {
                        r[r.length - 1] = lineCubic(p[0], c[3]);
                        continue;
                    }
                }
            }
            r.push(c);
        }
        if (r.length > 2) {
            var f = r[0], l = r[r.length - 1];
            if (isLinear(f) && isLinear(l)) {
                var e1 = norm(sub(l[3], l[0])), e2 = norm(sub(f[3], f[0]));
                if (dot(e1, e2) > 0.99999) {
                    r[0] = lineCubic(l[0], f[3]);
                    r.pop();
                }
            }
        }
        return r;
    }
}
