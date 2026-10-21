/*
@METADATA
{
  "name": "Color Key Creator",
  "description": "Create A Prepositioned Color Key On Proofs",
  "version": "1.10.1",
  "target": "illustrator",
  "tags": ["proof", "color", "key"]
}
@END_METADATA
*/

if (app.documents.length == 0) {
    alert("Please open a document first.");
} else {
    var doc = app.activeDocument;
    
    try {
        // Check if there is a selection
        if (doc.selection.length == 0) {
            alert("Please select some objects first.");
        } else {
        
            // === COLOR ANALYSIS ===
            var usedColors = {};
            var MAX_COLORS = 15;   // Warn the user above this many colors
            var SCAN_CAP = 100;    // Stop scanning entirely past this many (art is not key-able)
            var colorCount = 0;
            var scanAborted = false;

            function countUsedColors() {
                var n = 0;
                for (var key in usedColors) { n++; }
                return n;
            }

            // === PANTONE FLAGGING ===
            var PANTONE_HEADING = "Pantones needed for\raccurate color matching";
            var PANTONE_BODY = "If color is critical, now or in the\r" +
                               "future, please provide a pantone\r" +
                               "color for matching. Otherwise,\r" +
                               "colors will print as CMYK and\r" +
                               "cannot be color matched.";
            // Eyedropper icon outlines on a 22 x 22 pt grid. Point = [x, y, leftX, leftY, rightX, rightY]
            var PANTONE_ICON = [[[9.6,22.0,9.6,22.0,9.0,21.9],[7.8,21.6,8.4,21.7,6.4,21.2],[4.0,19.5,5.1,20.5,4.9,18.6],[6.5,17.0,5.7,17.8,7.7,17.5],[10.2,18.5,8.9,18.0,10.2,18.5],[10.2,22.0,10.2,22.0,10.2,22.0],[9.7,22.0,9.7,22.0,9.7,22.0]],[[11.9,22.0,11.9,22.0,11.9,22.0],[11.9,18.5,11.9,18.5,13.2,18.0],[15.6,17.0,14.4,17.5,16.4,17.8],[18.1,19.5,17.2,18.6,17.9,19.6],[17.5,20.0,17.7,19.8,16.1,21.0],[12.8,21.9,14.5,21.7,12.7,21.9],[12.5,22.0,12.6,21.9,12.5,22.0],[12.0,22.0,12.0,22.0,12.0,22.0]],[[22.0,12.4,22.0,12.4,21.9,12.9],[21.7,13.8,21.8,13.4,21.3,15.4],[19.5,18.0,20.5,16.8,18.6,17.1],[16.9,15.4,17.8,16.3,17.6,14.5],[18.3,12.1,18.1,13.4,18.3,12.0],[18.8,11.8,18.6,11.8,19.9,11.8],[22.0,11.8,21.0,11.8,22.0,11.8],[22.0,12.3,22.0,12.3,22.0,12.3]],[[2.5,3.9,2.5,3.9,3.4,4.8],[5.1,6.6,4.3,5.7,4.4,7.5],[3.7,9.8,3.9,8.6,3.7,9.8],[3.4,10.1,3.5,10.0,2.3,10.1],[-0.1,10.1,1.1,10.1,0.1,7.7],[2.4,3.9,1.0,5.7,2.4,3.9]],[[0.0,11.9,0.0,11.9,0.0,11.9],[3.6,11.9,3.6,11.9,4.1,13.1],[5.1,15.5,4.6,14.3,4.3,16.3],[2.6,18.0,3.5,17.1,1.1,16.3],[0.1,11.9,0.3,14.2,0.1,11.9]],[[4.0,2.5,4.0,2.5,5.7,1.1],[10.1,0.0,7.7,0.2,10.1,0.0],[10.1,3.5,10.1,3.5,8.9,4.0],[6.6,5.0,7.7,4.5,5.8,4.2],[4.1,2.5,4.9,3.4,4.1,2.5]],[[21.9,1.9,21.9,1.9,21.7,1.0],[20.3,0.1,21.2,0.4,19.4,-0.2],[18.0,0.7,18.6,0.1,17.5,1.2],[16.4,2.3,16.9,1.8,15.9,2.8],[14.9,3.9,15.4,3.3,14.6,3.5],[14.1,3.0,14.3,3.3,13.6,2.6],[12.5,3.0,13.0,2.6,12.1,3.4],[12.5,4.6,12.0,4.1,12.7,4.9],[13.4,5.3,13.1,5.1,12.0,6.7],[9.4,9.3,10.7,8.0,8.8,9.9],[8.5,11.5,8.4,10.6,8.5,12.0],[8.6,12.9,8.5,12.4,8.6,13.2],[9.1,13.3,8.8,13.3,9.3,13.3],[9.6,13.3,9.4,13.3,10.9,13.5],[12.9,12.3,11.9,13.2,14.0,11.2],[16.1,9.1,15.0,10.1,16.3,8.9],[16.6,8.4,16.4,8.7,17.0,8.8],[17.4,9.3,17.2,9.0,17.9,9.8],[19.1,9.3,18.6,9.8,19.6,8.8],[19.1,7.6,19.6,8.2,18.9,7.4],[18.3,6.9,18.6,7.1,19.2,6.0],[21.0,4.3,20.1,5.1,21.6,3.7],[22.3,2.3,22.1,3.1,22.3,2.3],[22.3,1.9,22.3,1.9,22.3,1.9],[22.3,1.7,22.3,1.8,22.3,1.7]],[[11.6,12.0,11.6,12.0,10.9,12.7],[9.6,11.7,10.0,12.5,9.6,11.6],[9.6,11.3,9.6,11.4,9.7,11.0],[10.1,10.4,9.8,10.7,11.4,9.1],[14.0,6.4,12.7,7.8,14.6,6.9],[15.6,8.0,15.1,7.5,14.3,9.3],[11.6,12.0,13.0,10.7,11.6,12.0]]];

            // True for colors that can't be color matched without a Pantone
            // Never flagged: true black/white, non-colors, and production marks (CutContour/Spot1 get the red warning)
            var NEVER_FLAG = ["CMYK Black", "CMYK White", "No Color", "[Registration]", "CutContour", "Spot1"];

            // Every color that isn't a Pantone gets flagged - process, RGB, LAB, gray,
            // gradients, patterns, and non-Pantone spots (Sherwin Williams etc.)
            function needsPantone(name) {
                for (var nf = 0; nf < NEVER_FLAG.length; nf++) {
                    if (name === NEVER_FLAG[nf]) return false;
                }
                if (/^(PANTONE|PMS)\b/i.test(name)) return false;
                if (/vinyl/i.test(name)) return false; // vinyl isn't printed - listed for callout only
                return true;
            }

            function makeOrange() {
                var c = new CMYKColor();
                c.cyan = 0; c.magenta = 55; c.yellow = 100; c.black = 0;
                return c;
            }

            function getFont(names) {
                for (var fn = 0; fn < names.length; fn++) {
                    try { return app.textFonts.getByName(names[fn]); } catch (eFn) {}
                }
                return null;
            }

            function getOrAddCharStyle(name) {
                try { return doc.characterStyles.getByName(name); } catch (eS) {}
                return doc.characterStyles.add(name);
            }

            // Draw the eyedropper icon with its top-left corner at (left, top)
            function drawPantoneIcon(left, top, size) {
                var sc = size / 22;
                var fill = makeOrange();
                var cp = doc.compoundPathItems.add();
                for (var si = 0; si < PANTONE_ICON.length; si++) {
                    var pts = PANTONE_ICON[si];
                    var anchors = [];
                    for (var pi = 0; pi < pts.length; pi++) {
                        anchors.push([left + pts[pi][0] * sc, top - pts[pi][1] * sc]);
                    }
                    var p = cp.pathItems.add();
                    p.setEntirePath(anchors);
                    p.closed = true;
                    for (var pj = 0; pj < pts.length; pj++) {
                        var pp = p.pathPoints[pj];
                        pp.leftDirection = [left + pts[pj][2] * sc, top - pts[pj][3] * sc];
                        pp.rightDirection = [left + pts[pj][4] * sc, top - pts[pj][5] * sc];
                    }
                    p.filled = true;
                    p.fillColor = fill;
                    p.stroked = false;
                    try { p.evenodd = true; } catch (eEo) {}
                }
                cp.name = "Pantone Flag";
                return cp;
            }
            
            // Function to add color to used colors list
            function addUsedColor(color) {
                try {
                    // Check if color object exists and has typename
                    if (!color || typeof color !== 'object' || !color.typename) {
                        var unknownName = "Error Reading Color: color object invalid";
                        if (!usedColors[unknownName]) {
                            usedColors[unknownName] = " (Error)";
                        }
                        return;
                    }
                    
                    // Debug: Log the color type we're examining
                    // alert("Found color type: " + color.typename);
                    
                    if (color.typename == "SpotColor" && color.spot) {
                        // This is a true spot color - use the spot name
                        var colorName = color.spot.name;
                        var isGlobalProcess = false;
                        try { isGlobalProcess = (color.spot.colorType == ColorModel.PROCESS); } catch (eT) {}
                        if (!usedColors[colorName]) {
                            usedColors[colorName] = isGlobalProcess ? " (Process)" : " (Spot)";
                        }
                    } else if (color.typename == "CMYKColor") {
                        // This is a process color - could be named or unnamed
                        var c = Math.round(color.cyan || 0);
                        var m = Math.round(color.magenta || 0);
                        var y = Math.round(color.yellow || 0);
                        var k = Math.round(color.black || 0);
                        
                        // Debug: Log CMYK values
                        // alert("CMYK found: C:" + c + " M:" + m + " Y:" + y + " K:" + k);
                        
                        var colorName;
                        
                        // Check for special standard process colors first
                        if (c == 0 && m == 0 && y == 0 && k == 100) {
                            colorName = "CMYK Black";
                        } else if (c == 0 && m == 0 && y == 0 && k == 0) {
                            colorName = "CMYK White";
                        } else {
                            // Don't do any spot color matching - just treat as Process color
                            colorName = "CMYK (C:" + c + " M:" + m + " Y:" + y + " K:" + k + ")";
                        }
                        
                        if (!usedColors[colorName]) {
                            usedColors[colorName] = " (Process)";
                        }
                    } else if (color.typename == "RGBColor") {
                        // RGB Color
                        var r = Math.round(color.red || 0);
                        var g = Math.round(color.green || 0);
                        var b = Math.round(color.blue || 0);
                        var colorName = "RGB Color (R:" + r + " G:" + g + " B:" + b + ")";
                        if (!usedColors[colorName]) {
                            usedColors[colorName] = " (RGB)";
                        }
                    } else if (color.typename == "LabColor") {
                        // LAB Color
                        var l = Math.round(color.l || 0);
                        var a = Math.round(color.a || 0);
                        var b = Math.round(color.b || 0);
                        var colorName = "LAB Color (L:" + l + " A:" + a + " B:" + b + ")";
                        if (!usedColors[colorName]) {
                            usedColors[colorName] = " (LAB)";
                        }
                    } else if (color.typename == "GrayColor") {
                        // Grayscale Color
                        var gray = Math.round(color.gray || 0);
                        var colorName = "Gray Color (" + gray + "%)";
                        if (!usedColors[colorName]) {
                            usedColors[colorName] = " (Gray)";
                        }
                    } else if (color.typename == "GradientColor") {
                        // Gradient Color
                        var gradientName = (color.gradient && color.gradient.name) ? color.gradient.name : "Unknown Gradient";
                        var colorName = "Gradient: " + gradientName;
                        if (!usedColors[colorName]) {
                            usedColors[colorName] = " (Gradient)";
                        }
                    } else if (color.typename == "PatternColor") {
                        // Pattern Color
                        var patternName = (color.pattern && color.pattern.name) ? color.pattern.name : "Unknown Pattern";
                        var colorName = "Pattern: " + patternName;
                        if (!usedColors[colorName]) {
                            usedColors[colorName] = " (Pattern)";
                        }
                    } else if (color.typename == "NoColor") {
                        // No Color (transparent)
                        if (!usedColors["No Color"]) {
                            usedColors["No Color"] = " (None)";
                        }
                    } else {
                        // Debug: Log unknown color types
                        var unknownName = "Unknown Color Type: " + (color.typename || "undefined");
                        if (!usedColors[unknownName]) {
                            usedColors[unknownName] = " (Unknown)";
                        }
                    }
                } catch (e) {
                    // If we can't read the color, add it as unknown
                    var unknownName = "Error Reading Color: " + e.toString();
                    if (!usedColors[unknownName]) {
                        usedColors[unknownName] = "";
                    }
                }
            }
            
            // Function to recursively check page items for used colors
            function checkItemsForColors(items) {
                for (var i = 0; i < items.length; i++) {
                    if (scanAborted) return;
                    var item = items[i];
                    
                    try {
                        // Check fill color
                        if (item.filled && item.fillColor) {
                            addUsedColor(item.fillColor);
                        }
                        
                        // Check stroke color
                        if (item.stroked && item.strokeColor) {
                            addUsedColor(item.strokeColor);
                        }

                        // Bail out early on art with an unreasonable number of colors
                        colorCount = countUsedColors();
                        if (colorCount > SCAN_CAP) {
                            scanAborted = true;
                            return;
                        }
                        
                        // Recursively process grouped items
                        if (item.typename == "GroupItem" && item.pageItems.length > 0) {
                            checkItemsForColors(item.pageItems);
                        }
                        
                        // Check compound path items
                        if (item.typename == "CompoundPathItem" && item.pathItems.length > 0) {
                            checkItemsForColors(item.pathItems);
                        }
                        
                    } catch (e) {
                        // Ignore errors for items without color properties
                    }
                }
            }
            
            // Check only selected items for used colors
            checkItemsForColors(doc.selection);
            
            // Create a lookup table for spot colors to improve performance
            var spotColorLookup = {};
            for (var s = 0; s < doc.spots.length; s++) {
                try {
                    var spot = doc.spots[s];
                    if (spot.color && spot.color.typename === "CMYKColor") {
                        var spotC = Math.round(spot.color.cyan);
                        var spotM = Math.round(spot.color.magenta);
                        var spotY = Math.round(spot.color.yellow);
                        var spotK = Math.round(spot.color.black);
                        var cmykKey = spotC + "," + spotM + "," + spotY + "," + spotK;
                        spotColorLookup[cmykKey] = {
                            name: spot.name,
                            color: spot.color
                        };
                    }
                } catch (e) {
                    // Skip this spot if error
                }
            }
            
            // Collect and sort color names
            var colorNames = [];
            for (var colorName in usedColors) {
                colorNames.push(colorName);
            }
            colorNames.sort();
            
            var proceed = true;
            if (scanAborted) {
                alert("The selection contains more than " + SCAN_CAP + " colors.\n\n" +
                      "This art is too complex for a color key (gradients, raster traces, or photo-based art are common causes). " +
                      "No key was created.", "Color Key - Too Many Colors");
                proceed = false;
            } else if (colorNames.length > MAX_COLORS) {
                proceed = confirm("The selection contains " + colorNames.length + " colors (limit is " + MAX_COLORS + ").\n\n" +
                                  "Building a key this large can take a long time or crash Illustrator.\n\n" +
                                  "Continue anyway?", true, "Color Key - Too Many Colors");
            }

            if (!proceed) {
                // User cancelled or art was too complex - do nothing
            } else if (colorNames.length == 0) {
                alert("No colors found in the selected objects.");
            } else {
                
                // === PANTONE CHECK ===
                var flagged = {};
                var flaggedList = [];
                for (var fi = 0; fi < colorNames.length; fi++) {
                    if (needsPantone(colorNames[fi])) {
                        flagged[colorNames[fi]] = true;
                        flaggedList.push(colorNames[fi]);
                    }
                }
                var flagPantone = false;
                if (flaggedList.length > 0) {
                    flagPantone = confirm(flaggedList.length + " color(s) are not Pantone (PMS):\n\n  " +
                                          flaggedList.join("\n  ") +
                                          "\n\nFlag them and add the \"Pantones needed\" disclaimer?",
                                          false, "Color Key - Pantone Check");
                }
                if (flagPantone) {
                    // List flagged colors first so they read as a block
                    var unflagged = [];
                    for (var ui = 0; ui < colorNames.length; ui++) {
                        if (!flagged[colorNames[ui]]) unflagged.push(colorNames[ui]);
                    }
                    colorNames = flaggedList.concat(unflagged);
                }

                // === CREATE COLOR LABELS (auto-fit into the key area) ===
                var artboard = doc.artboards[doc.artboards.getActiveArtboardIndex()];
                var artboardBounds = artboard.artboardRect;
                var artboardLeft = artboardBounds[0];
                var artboardTop = artboardBounds[1];
                var startX = artboardLeft + (6.68 * 72); // 6.68" from left edge of artboard
                var startY = artboardTop - (7.145 * 72); // 7.145" down from top edge of artboard

                // Key area on the proof template
                var areaLeft = startX + 5;               // left edge of first swatch
                var areaRight = artboardBounds[2] - 9;   // 1/8" inside the artboard edge
                var areaTop = startY + 2.5;              // top of first swatch
                var areaHeight = 59.4;                   // original 5 rows at 12.6pt pitch
                var BASE_PITCH = 12.6;                   // row pitch at full size
                var BASE_CIRCLE = 9;                     // swatch diameter at full size
                var BASE_GUTTER = 8;                     // space between columns
                var BASE_DISC_GAP = 10;                  // space between key and disclaimer
                var MIN_SCALE = 0.8;                     // never shrink the key below 80% (5.6pt text)
                var FLAG_MARKER = "ring";                // "ring" = dashed orange ring around the swatch, "icon" = eyedropper after the name
                var RING_SIZE = 12;                      // flag ring diameter at full size
                var CAP_HEIGHT = 0.716;                  // Arial cap height (fraction of font size)
                var SHORT_NAMES = true;                  // "c0 m16 y100 k58" instead of "CMYK (C:0 M:16 Y:100 K:58)"

                function displayName(name) {
                    if (/^PANTONE\s+/i.test(name)) return name.replace(/^PANTONE\s+/i, "PMS ");
                    if (!SHORT_NAMES) return name;
                    var m = name.match(/^CMYK \(C:(\d+) M:(\d+) Y:(\d+) K:(\d+)\)$/);
                    if (m) return "c" + m[1] + " m" + m[2] + " y" + m[3] + " k" + m[4];
                    m = name.match(/^RGB Color \(R:(\d+) G:(\d+) B:(\d+)\)$/);
                    if (m) return "r" + m[1] + " g" + m[2] + " b" + m[3];
                    m = name.match(/^LAB Color \(L:(-?\d+) A:(-?\d+) B:(-?\d+)\)$/);
                    if (m) return "L" + m[1] + " a" + m[2] + " b" + m[3];
                    m = name.match(/^Gray Color \((\d+)%\)$/);
                    if (m) return "Gray " + m[1] + "%";
                    return name;
                }

                // Store original selection to restore later
                var originalSelection = doc.selection;
                doc.selection = null;

                // Pre-cache fonts for reliable text creation
                var regularFont = getFont(["ArialMT"]);
                var boldFont = getFont(["Arial-BoldMT", "ArialMT-Bold"]);

                // Character styles force consistent text attributes
                var normalStyle = getOrAddCharStyle("_ColorKeyNormal");
                normalStyle.characterAttributes.size = 7;
                var normalFill = new CMYKColor();
                normalFill.cyan = 0; normalFill.magenta = 0; normalFill.yellow = 0; normalFill.black = 100;
                normalStyle.characterAttributes.fillColor = normalFill;
                if (regularFont) normalStyle.characterAttributes.textFont = regularFont;

                var warningStyle = getOrAddCharStyle("_ColorKeyWarning");
                warningStyle.characterAttributes.size = 7;
                var warningFill = new CMYKColor();
                warningFill.cyan = 0; warningFill.magenta = 100; warningFill.yellow = 100; warningFill.black = 0;
                warningStyle.characterAttributes.fillColor = warningFill;
                if (boldFont) warningStyle.characterAttributes.textFont = boldFont;

                var flagHeadStyle = null;
                var flagBodyStyle = null;
                if (flagPantone) {
                    flagHeadStyle = getOrAddCharStyle("_ColorKeyFlagHead");
                    flagHeadStyle.characterAttributes.size = 7;
                    flagHeadStyle.characterAttributes.autoLeading = false;
                    flagHeadStyle.characterAttributes.leading = 8;
                    flagHeadStyle.characterAttributes.fillColor = makeOrange();
                    var headFont = getFont(["MyriadPro-Black", "MyriadPro-Bold", "Arial-BoldMT", "ArialMT-Bold"]);
                    if (headFont) flagHeadStyle.characterAttributes.textFont = headFont;

                    flagBodyStyle = getOrAddCharStyle("_ColorKeyFlagBody");
                    flagBodyStyle.characterAttributes.size = 6;
                    flagBodyStyle.characterAttributes.autoLeading = false;
                    flagBodyStyle.characterAttributes.leading = 7.2;
                    flagBodyStyle.characterAttributes.fillColor = normalFill;
                    var bodyFont = getFont(["MyriadPro-Regular", "ArialMT"]);
                    if (bodyFont) flagBodyStyle.characterAttributes.textFont = bodyFont;
                }

                // Group via the menu command; if Illustrator drops it (it can be flaky in loops),
                // fall back to building the group directly
                function groupItems(items, name) {
                    var g = null;
                    try {
                        doc.selection = null;
                        doc.selection = items;
                        app.executeMenuCommand("group");
                        if (doc.selection.length == 1 && doc.selection[0].typename == "GroupItem") {
                            g = doc.selection[0];
                        }
                    } catch (eG) {
                        g = null;
                    }
                    if (!g) {
                        try {
                            g = items[0].parent.groupItems.add();
                            for (var gi = items.length - 1; gi >= 0; gi--) {
                                items[gi].move(g, ElementPlacement.PLACEATBEGINNING);
                            }
                        } catch (eG2) {
                            g = null;
                        }
                    }
                    try { doc.selection = null; } catch (eG3) {}
                    if (g) {
                        try { g.name = name; } catch (eG4) {}
                    }
                    return g;
                }

                // --- Phase 1: build every entry at full size, stacked at the first row ---
                // entry = { group, width, dy } where dy = swatch center minus group top
                var entries = [];
                var failedEntries = [];
                var rowCenterY = startY - 2;

                for (var i = 0; i < colorNames.length; i++) {
                    var colorName = colorNames[i];

                    var isCutContourProcess = (colorName === "CutContour" && usedColors[colorName] === " (Process)");
                    var isSpot1Process = (colorName === "Spot1" && usedColors[colorName] === " (Process)");
                    var isWarningColor = isCutContourProcess || isSpot1Process;

                    // --- Text ---
                    var textFrame = null;
                    try {
                        textFrame = doc.textFrames.add();
                        textFrame.contents = displayName(colorName);
                        (isWarningColor ? warningStyle : normalStyle).applyTo(textFrame.textRange);
                        textFrame.left = areaLeft + BASE_CIRCLE + 2.44;
                        // Put the baseline so capitals are centered on the swatch
                        var targetBaseline = rowCenterY - (CAP_HEIGHT * 7) / 2;
                        try {
                            textFrame.translate(0, targetBaseline - textFrame.anchor[1]);
                        } catch (eA) {
                            textFrame.top = rowCenterY + 5.5;
                        }
                    } catch (e) {
                        textFrame = null;
                    }

                    // --- Swatch circle ---
                    var colorCircle = null;
                    try {
                        var circleSize = BASE_CIRCLE;
                        colorCircle = doc.pathItems.ellipse(rowCenterY + circleSize / 2, areaLeft, circleSize, circleSize);

                        var circleColor;
                        if (doc.documentColorSpace == DocumentColorSpace.RGB) {
                            circleColor = new RGBColor();
                            if (colorName.indexOf("CMYK (C:") === 0) {
                                var cmykMatch = colorName.match(/C:(\d+) M:(\d+) Y:(\d+) K:(\d+)/);
                                if (cmykMatch) {
                                    var c = parseInt(cmykMatch[1]) / 100;
                                    var m = parseInt(cmykMatch[2]) / 100;
                                    var y = parseInt(cmykMatch[3]) / 100;
                                    var k = parseInt(cmykMatch[4]) / 100;
                                    circleColor.red = Math.round(255 * (1 - c) * (1 - k));
                                    circleColor.green = Math.round(255 * (1 - m) * (1 - k));
                                    circleColor.blue = Math.round(255 * (1 - y) * (1 - k));
                                }
                            } else if (colorName.indexOf("RGB Color (R:") === 0) {
                                var rgbMatch = colorName.match(/R:(\d+) G:(\d+) B:(\d+)/);
                                if (rgbMatch) {
                                    circleColor.red = parseInt(rgbMatch[1]);
                                    circleColor.green = parseInt(rgbMatch[2]);
                                    circleColor.blue = parseInt(rgbMatch[3]);
                                }
                            } else if (colorName === "CMYK Black") {
                                circleColor.red = 0; circleColor.green = 0; circleColor.blue = 0;
                            } else if (colorName === "CMYK White") {
                                circleColor.red = 255; circleColor.green = 255; circleColor.blue = 255;
                            } else if (colorName.indexOf("Gray Color (") === 0) {
                                var grayMatch = colorName.match(/(\d+)%/);
                                if (grayMatch) {
                                    var rgbGray = Math.round(255 - parseInt(grayMatch[1]) * 2.55);
                                    circleColor.red = rgbGray; circleColor.green = rgbGray; circleColor.blue = rgbGray;
                                }
                            } else {
                                circleColor.red = 128; circleColor.green = 128; circleColor.blue = 128;
                            }
                        } else {
                            circleColor = new CMYKColor();
                            if (colorName.indexOf("CMYK (C:") === 0) {
                                var cmykMatch = colorName.match(/C:(\d+) M:(\d+) Y:(\d+) K:(\d+)/);
                                if (cmykMatch) {
                                    circleColor.cyan = parseInt(cmykMatch[1]);
                                    circleColor.magenta = parseInt(cmykMatch[2]);
                                    circleColor.yellow = parseInt(cmykMatch[3]);
                                    circleColor.black = parseInt(cmykMatch[4]);
                                }
                            } else if (colorName.indexOf("RGB Color (R:") === 0) {
                                var rgbMatch = colorName.match(/R:(\d+) G:(\d+) B:(\d+)/);
                                if (rgbMatch) {
                                    var r = parseInt(rgbMatch[1]);
                                    var g = parseInt(rgbMatch[2]);
                                    var b = parseInt(rgbMatch[3]);
                                    var rP = r / 255, gP = g / 255, bP = b / 255;
                                    var k = 1 - Math.max(rP, Math.max(gP, bP));
                                    circleColor.cyan = (k < 1) ? Math.round((1 - rP - k) / (1 - k) * 100) : 0;
                                    circleColor.magenta = (k < 1) ? Math.round((1 - gP - k) / (1 - k) * 100) : 0;
                                    circleColor.yellow = (k < 1) ? Math.round((1 - bP - k) / (1 - k) * 100) : 0;
                                    circleColor.black = Math.round(k * 100);
                                }
                            } else if (colorName.indexOf("Gray Color (") === 0) {
                                var grayMatch = colorName.match(/(\d+)%/);
                                if (grayMatch) {
                                    circleColor.cyan = 0; circleColor.magenta = 0; circleColor.yellow = 0;
                                    circleColor.black = parseInt(grayMatch[1]);
                                }
                            } else if (colorName === "CMYK Black") {
                                circleColor.cyan = 0; circleColor.magenta = 0; circleColor.yellow = 0; circleColor.black = 100;
                            } else if (colorName === "CMYK White" || colorName === "No Color") {
                                circleColor.cyan = 0; circleColor.magenta = 0; circleColor.yellow = 0; circleColor.black = 0;
                            } else if (colorName.indexOf("Gradient:") === 0 || colorName.indexOf("Pattern:") === 0 || colorName.indexOf("LAB Color (") === 0) {
                                circleColor.cyan = 0; circleColor.magenta = 0; circleColor.yellow = 0; circleColor.black = 50;
                            } else {
                                var spotFound = false;
                                for (var lookupKey in spotColorLookup) {
                                    if (spotColorLookup[lookupKey].name === colorName) {
                                        var spotData = spotColorLookup[lookupKey];
                                        circleColor.cyan = spotData.color.cyan;
                                        circleColor.magenta = spotData.color.magenta;
                                        circleColor.yellow = spotData.color.yellow;
                                        circleColor.black = spotData.color.black;
                                        spotFound = true;
                                        break;
                                    }
                                }
                                if (!spotFound) {
                                    circleColor.cyan = 50; circleColor.magenta = 50; circleColor.yellow = 50; circleColor.black = 50;
                                }
                            }
                        }

                        colorCircle.fillColor = circleColor;
                        colorCircle.stroked = true;
                        var strokeColor = new CMYKColor();
                        strokeColor.cyan = 0; strokeColor.magenta = 0; strokeColor.yellow = 0; strokeColor.black = 100;
                        colorCircle.strokeColor = strokeColor;
                        colorCircle.strokeWidth = 0.5;
                    } catch (e) {
                        colorCircle = null;
                    }

                    // --- Pantone flag icon after the color name ---
                    var flagIcon = null;
                    if (flagPantone && flagged[colorName] && FLAG_MARKER === "ring") {
                        try {
                            var ringR = RING_SIZE / 2;
                            flagIcon = doc.pathItems.ellipse(rowCenterY + ringR, areaLeft + BASE_CIRCLE / 2 - ringR, RING_SIZE, RING_SIZE);
                            flagIcon.filled = false;
                            flagIcon.stroked = true;
                            flagIcon.strokeColor = makeOrange();
                            flagIcon.strokeWidth = 0.9;
                            var seg = (Math.PI * RING_SIZE) / 8; // 8 dashes like the icon's ring
                            flagIcon.strokeDashes = [seg * 0.68, seg * 0.32];
                            flagIcon.name = "Pantone Flag";
                        } catch (eR) {
                            flagIcon = null;
                        }
                    } else if (flagPantone && flagged[colorName]) {
                        try {
                            var flagSize = 7;
                            var flagLeft = textFrame ? textFrame.geometricBounds[2] + 3 : areaLeft + 60;
                            flagIcon = drawPantoneIcon(flagLeft, rowCenterY + flagSize / 2, flagSize);
                        } catch (eF) {
                            flagIcon = null;
                        }
                    }

                    // --- Group the entry ---
                    try {
                        var itemsToGroup = [];
                        if (colorCircle) itemsToGroup.push(colorCircle);
                        if (textFrame) itemsToGroup.push(textFrame);
                        if (flagIcon) itemsToGroup.push(flagIcon);
                        if (itemsToGroup.length > 0) {
                            var eg = groupItems(itemsToGroup, colorName);
                            if (eg) {
                                var egb = eg.geometricBounds; // [left, top, right, bottom]
                                entries.push({ group: eg, left: areaLeft - egb[0], right: egb[2] - areaLeft, dy: rowCenterY - egb[1] });
                            } else {
                                failedEntries.push(colorName);
                            }
                        } else {
                            failedEntries.push(colorName);
                        }
                    } catch (eg2) {
                        failedEntries.push(colorName + " (" + eg2.toString() + ")");
                    }
                }

                // --- Pantone disclaimer block (compact, stacked) ---
                var disclaimer = null;
                if (flagPantone && entries.length > 0) {
                    try {
                        var bigIconSize = 16;
                        var dItems = [];
                        dItems.push(drawPantoneIcon(areaLeft, areaTop, bigIconSize));

                        var headFrame = doc.textFrames.add();
                        headFrame.contents = PANTONE_HEADING;
                        flagHeadStyle.applyTo(headFrame.textRange);
                        headFrame.left = areaLeft + bigIconSize + 3;
                        headFrame.top = areaTop;
                        dItems.push(headFrame);

                        var bodyFrame = doc.textFrames.add();
                        bodyFrame.contents = PANTONE_BODY;
                        flagBodyStyle.applyTo(bodyFrame.textRange);
                        bodyFrame.left = headFrame.left;
                        bodyFrame.top = headFrame.top - headFrame.height - 0.5;
                        dItems.push(bodyFrame);

                        var dg = groupItems(dItems, "Pantone Disclaimer");
                        if (dg) {
                            var dgb = dg.geometricBounds;
                            disclaimer = { group: dg, width: dgb[2] - dgb[0] };
                        }
                    } catch (eD) {}
                }

                // --- Phase 2: find the largest scale where everything fits ---
                function layoutFor(s) {
                    var pitch = BASE_PITCH * s;
                    var d = BASE_CIRCLE * s;
                    var rows = Math.max(1, Math.floor((areaHeight - d) / pitch + 0.001) + 1);
                    var cols = Math.ceil(entries.length / rows);
                    var colWidths = [];
                    var total = 0;
                    for (var c = 0; c < cols; c++) {
                        var w = 0;
                        for (var r = 0; r < rows; r++) {
                            var idx = c * rows + r;
                            if (idx < entries.length) w = Math.max(w, (maxLeft + entries[idx].right) * s);
                        }
                        colWidths.push(w);
                        total += w;
                    }
                    total += (cols - 1) * BASE_GUTTER * s;
                    if (disclaimer) total += BASE_DISC_GAP * s + disclaimer.width; // disclaimer is never scaled
                    return { s: s, pitch: pitch, d: d, rows: rows, cols: cols, colWidths: colWidths, total: total };
                }

                // Widest thing hanging left of a swatch (the flag ring) - keeps swatches aligned
                var maxLeft = 0;
                for (var ml = 0; ml < entries.length; ml++) maxLeft = Math.max(maxLeft, entries[ml].left);
                var availWidth = areaRight - (areaLeft - maxLeft);
                var layout = null;
                for (var step = 0; step <= 30; step++) {
                    var sTry = Math.round((1 - step * 0.025) * 1000) / 1000;
                    if (sTry < MIN_SCALE) break;
                    layout = layoutFor(sTry);
                    if (layout.total <= availWidth) break;
                }
                var keyOverflow = layout && layout.total > availWidth;

                // --- Phase 3: scale and place ---
                var entryGroups = [];
                if (layout) {
                    var pct = layout.s * 100;
                    var x = areaLeft - maxLeft; // left edge of column 1
                    for (var c = 0; c < layout.cols; c++) {
                        for (var r = 0; r < layout.rows; r++) {
                            var idx = c * layout.rows + r;
                            if (idx >= entries.length) break;
                            var en = entries[idx];
                            try {
                                if (pct < 100) en.group.resize(pct, pct, true, true, true, true, pct, Transformation.TOPLEFT);
                                var gb = en.group.geometricBounds;
                                var targetCenter = areaTop - layout.d / 2 - r * layout.pitch;
                                var currentCenter = gb[1] + en.dy * layout.s;
                                var swatchLeft = x + maxLeft * layout.s;
                                var currentSwatchLeft = gb[0] + en.left * layout.s;
                                en.group.translate(swatchLeft - currentSwatchLeft, targetCenter - currentCenter);
                                entryGroups.push(en.group);
                            } catch (eP) {
                                var enName = "";
                                try { enName = en.group.name; } catch (eN) {}
                                failedEntries.push(enName + " (" + eP.toString() + ")");
                            }
                        }
                        x += layout.colWidths[c] + BASE_GUTTER * layout.s;
                    }

                    if (disclaimer) {
                        try {
                            var dgb2 = disclaimer.group.geometricBounds;
                            var dLeft = x - BASE_GUTTER * layout.s + BASE_DISC_GAP * layout.s;
                            disclaimer.group.translate(dLeft - dgb2[0], areaTop - dgb2[1]);
                            entryGroups.push(disclaimer.group);
                        } catch (eP2) {}
                    }
                }

                // Group all entry subgroups into a parent "Color Labels" group
                if (entryGroups.length > 0) {
                    try {
                        doc.selection = null;
                        doc.selection = entryGroups;
                        app.executeMenuCommand("group");
                        if (doc.selection.length == 1) {
                            doc.selection[0].name = "Color Labels";
                        }
                        doc.selection = null;
                    } catch (e) {}
                }

                // Clean up temporary character styles
                try { normalStyle.remove(); } catch (e) {}
                try { warningStyle.remove(); } catch (e) {}
                try { if (flagHeadStyle) flagHeadStyle.remove(); } catch (e) {}
                try { if (flagBodyStyle) flagBodyStyle.remove(); } catch (e) {}

                // Restore original selection
                doc.selection = originalSelection;

                if (failedEntries.length > 0) {
                    alert("These colors couldn't be placed in the key layout:\n\n  " +
                          failedEntries.join("\n  ") + "\n\nPlease move them by hand.", "Color Key");
                }

                if (keyOverflow) {
                    alert("Too many colors to fit the key area, even at " + Math.round(MIN_SCALE * 100) + "% size.\n\nThe key runs past the right edge - check the layout.");
                }
                
            } // End of colorNames.length > 0 check
        } // End of selection.length > 0 check
        
    } catch (error) {
        alert("Error creating color labels: " + error.toString());
    }
}