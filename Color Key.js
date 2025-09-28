/*
@METADATA
{
  "name": "Color Key Creator",
  "description": "Create A Prepositioned Color Key On Proofs",
  "version": "1.0",
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
                        if (!usedColors[colorName]) {
                            usedColors[colorName] = " (Spot)";
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
            
            if (colorNames.length == 0) {
                alert("No colors found in the selected objects.");
            } else {
                
                // === CREATE COLOR LABELS AT SPECIFIED COORDINATES ===
                // Get artboard bounds to calculate proper coordinates
                var artboard = doc.artboards[doc.artboards.getActiveArtboardIndex()];
                var artboardBounds = artboard.artboardRect;
                var artboardLeft = artboardBounds[0]; // Left X coordinate
                var artboardTop = artboardBounds[1]; // Top Y coordinate
                var startX = artboardLeft + (10 * 72); // 10 inches from left edge of artboard
                var startY = artboardTop - (8.75 * 72); // 8.75 inches down from top edge of artboard
                var yPosition = startY;
                var lineSpacing = 23; // Space between color entries
                
                // Store original selection to restore later
                var originalSelection = doc.selection;
                doc.selection = null;
                
                // Create all objects and collect them in an array
                var createdObjects = [];
                
                // Pre-cache font for better performance
                var boldFont = null;
                try {
                    boldFont = app.textFonts.getByName("ArialMT-Bold");
                } catch (e) {
                    // Bold font not available
                }
                
                // Helper function to add text
                function addText(text, x, y, fontSize, isWarning) {
                    try {
                        var textFrame = doc.textFrames.add();
                        textFrame.contents = text;
                        textFrame.left = x;
                        textFrame.top = y;
                        textFrame.textRange.characterAttributes.size = fontSize;
                        
                        // Set color based on warning status
                        if (isWarning) {
                            // Red color for warnings
                            var redColor = new CMYKColor();
                            redColor.cyan = 0;
                            redColor.magenta = 100;
                            redColor.yellow = 100;
                            redColor.black = 0;
                            textFrame.textRange.characterAttributes.fillColor = redColor;
                            
                            // Use cached bold font
                            if (boldFont) {
                                textFrame.textRange.characterAttributes.textFont = boldFont;
                            }
                        } else {
                            // Black color for regular text
                            var blackColor = new CMYKColor();
                            blackColor.cyan = 0;
                            blackColor.magenta = 0;
                            blackColor.yellow = 0;
                            blackColor.black = 100;
                            textFrame.textRange.characterAttributes.fillColor = blackColor;
                        }
                        
                        return textFrame;
                    } catch (e) {
                        // If text creation fails, return null
                        return null;
                    }
                }
                
                // Add each color with swatch
                for (var i = 0; i < colorNames.length; i++) {
                    var colorName = colorNames[i];
                    
                    // Check if this is a warning color (CutContour or Spot1 used as Process)
                    var isCutContourProcess = (colorName === "CutContour" && usedColors[colorName] === " (Process)");
                    var isSpot1Process = (colorName === "Spot1" && usedColors[colorName] === " (Process)");
                    var isWarningColor = isCutContourProcess || isSpot1Process;
                    
                    // Add the color name text (offset to make room for color circle)
                    var textFrame = addText(colorName, startX + 40, yPosition, 11, isWarningColor);
                    if (textFrame) {
                        createdObjects.push(textFrame);
                    }
                    
                    // Create color swatch circle
                    try {
                        var circleSize = 15;
                        var circleLeft = startX + 20;
                        var circleTop = yPosition;
                        
                        // Create circle using ellipse
                        var colorCircle = doc.pathItems.ellipse(circleTop, circleLeft, circleSize, circleSize);
                        
                        // Determine the color for the circle based on document color mode
                        var circleColor;
                        
                        if (doc.documentColorSpace == DocumentColorSpace.RGB) {
                            // RGB Document - create RGB colors
                            circleColor = new RGBColor();
                            
                            if (colorName.indexOf("CMYK (C:") === 0) {
                                // Convert CMYK to RGB for display
                                var cmykMatch = colorName.match(/C:(\d+) M:(\d+) Y:(\d+) K:(\d+)/);
                                if (cmykMatch) {
                                    var c = parseInt(cmykMatch[1]) / 100;
                                    var m = parseInt(cmykMatch[2]) / 100;
                                    var y = parseInt(cmykMatch[3]) / 100;
                                    var k = parseInt(cmykMatch[4]) / 100;
                                    
                                    // Convert CMYK to RGB
                                    circleColor.red = Math.round(255 * (1 - c) * (1 - k));
                                    circleColor.green = Math.round(255 * (1 - m) * (1 - k));
                                    circleColor.blue = Math.round(255 * (1 - y) * (1 - k));
                                }
                            } else if (colorName.indexOf("RGB (R:") === 0) {
                                // Extract RGB values directly
                                var rgbMatch = colorName.match(/R:(\d+) G:(\d+) B:(\d+)/);
                                if (rgbMatch) {
                                    circleColor.red = parseInt(rgbMatch[1]);
                                    circleColor.green = parseInt(rgbMatch[2]);
                                    circleColor.blue = parseInt(rgbMatch[3]);
                                }
                            } else if (colorName === "CMYK Black") {
                                circleColor.red = 0;
                                circleColor.green = 0;
                                circleColor.blue = 0;
                            } else if (colorName === "CMYK White") {
                                circleColor.red = 255;
                                circleColor.green = 255;
                                circleColor.blue = 255;
                            } else if (colorName.indexOf("Gray (") === 0) {
                                var grayMatch = colorName.match(/(\d+)%/);
                                if (grayMatch) {
                                    var grayValue = parseInt(grayMatch[1]);
                                    var rgbGray = Math.round(grayValue * 2.55); // Convert percentage to 0-255
                                    circleColor.red = rgbGray;
                                    circleColor.green = rgbGray;
                                    circleColor.blue = rgbGray;
                                }
                            } else {
                                // For spot colors or unknown, try to get RGB equivalent or use gray
                                circleColor.red = 128;
                                circleColor.green = 128;
                                circleColor.blue = 128;
                            }
                            
                        } else {
                            // CMYK Document - create CMYK colors (existing logic)
                            circleColor = new CMYKColor();
                        
                            if (colorName.indexOf("CMYK (C:") === 0) {
                                // Extract CMYK values from process color name
                                var cmykMatch = colorName.match(/C:(\d+) M:(\d+) Y:(\d+) K:(\d+)/);
                                if (cmykMatch) {
                                    circleColor.cyan = parseInt(cmykMatch[1]);
                                    circleColor.magenta = parseInt(cmykMatch[2]);
                                    circleColor.yellow = parseInt(cmykMatch[3]);
                                    circleColor.black = parseInt(cmykMatch[4]);
                                }
                            } else if (colorName.indexOf("RGB (R:") === 0) {
                                // Convert RGB to CMYK approximation
                                var rgbMatch = colorName.match(/R:(\d+) G:(\d+) B:(\d+)/);
                                if (rgbMatch) {
                                    var r = parseInt(rgbMatch[1]);
                                    var g = parseInt(rgbMatch[2]);
                                    var b = parseInt(rgbMatch[3]);
                                    
                                    // Convert RGB to CMYK
                                    var rPercent = r / 255;
                                    var gPercent = g / 255;
                                    var bPercent = b / 255;
                                    
                                    var k = 1 - Math.max(rPercent, Math.max(gPercent, bPercent));
                                    var c = (k < 1) ? (1 - rPercent - k) / (1 - k) : 0;
                                    var m = (k < 1) ? (1 - gPercent - k) / (1 - k) : 0;
                                    var y = (k < 1) ? (1 - bPercent - k) / (1 - k) : 0;
                                    
                                    circleColor.cyan = Math.round(c * 100);
                                    circleColor.magenta = Math.round(m * 100);
                                    circleColor.yellow = Math.round(y * 100);
                                    circleColor.black = Math.round(k * 100);
                                }
                            } else if (colorName.indexOf("Gray (") === 0) {
                                // Extract gray percentage
                                var grayMatch = colorName.match(/(\d+)%/);
                                if (grayMatch) {
                                    var grayValue = parseInt(grayMatch[1]);
                                    circleColor.cyan = 0;
                                    circleColor.magenta = 0;
                                    circleColor.yellow = 0;
                                    circleColor.black = 100 - grayValue;
                                }
                            } else if (colorName === "CMYK Black") {
                                circleColor.cyan = 0;
                                circleColor.magenta = 0;
                                circleColor.yellow = 0;
                                circleColor.black = 100;
                            } else if (colorName === "CMYK White") {
                                circleColor.cyan = 0;
                                circleColor.magenta = 0;
                                circleColor.yellow = 0;
                                circleColor.black = 0;
                            } else if (colorName === "No Color") {
                                // White with dashed stroke to indicate transparency
                                circleColor.cyan = 0;
                                circleColor.magenta = 0;
                                circleColor.yellow = 0;
                                circleColor.black = 0;
                            } else if (colorName.indexOf("Gradient:") === 0 || colorName.indexOf("Pattern:") === 0 || colorName.indexOf("LAB (") === 0) {
                                // For complex colors, use a neutral gray
                                circleColor.cyan = 0;
                                circleColor.magenta = 0;
                                circleColor.yellow = 0;
                                circleColor.black = 50;
                            } else {
                                // For spot colors, use lookup table for better performance
                                var spotFound = false;
                                
                                // First try to find by name in our lookup
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
                                    // Default color for unknown spots or colors
                                    circleColor.cyan = 50;
                                    circleColor.magenta = 50;
                                    circleColor.yellow = 50;
                                    circleColor.black = 50;
                                }
                            }
                        }
                        
                        colorCircle.fillColor = circleColor;
                        
                        // Debug: Force a known RGB color to test if assignment works
                        // var testColor = new RGBColor();
                        // testColor.red = 255;
                        // testColor.green = 0;
                        // testColor.blue = 0;
                        // colorCircle.fillColor = testColor;
                        
                        // Add thin black stroke
                        colorCircle.stroked = true;
                        var strokeColor = new CMYKColor();
                        strokeColor.cyan = 0;
                        strokeColor.magenta = 0;
                        strokeColor.yellow = 0;
                        strokeColor.black = 100;
                        colorCircle.strokeColor = strokeColor;
                        colorCircle.strokeWidth = 0.75;
                        
                        // Align circle vertically to text center
                        if (textFrame && colorCircle) {
                            try {
                                // Get center point of text
                                var textCenterPoint = {
                                    "h": textFrame.left + textFrame.width / 2,
                                    "v": textFrame.top - textFrame.height / 2
                                };
                                
                                // Align circle vertically to text center
                                colorCircle.top = textCenterPoint.v + (colorCircle.height / 2);
                                
                            } catch (alignError) {
                                // If alignment calculation fails, continue without it
                            }
                        }
                        
                        createdObjects.push(colorCircle);
                        
                    } catch (e) {
                        // If circle creation fails, continue without it
                    }
                    
                    yPosition -= lineSpacing;
                }
                
                // Group all created objects
                if (createdObjects.length > 0) {
                    try {
                        // Clear any existing selection first
                        doc.selection = null;
                        // Select only the objects we created
                        doc.selection = createdObjects;
                        // Group them using the menu command
                        app.executeMenuCommand("group");
                        // Name the group if it was created successfully
                        if (doc.selection.length == 1 && doc.selection[0].typename == "GroupItem") {
                            doc.selection[0].name = "Color Labels";
                        }
                        // Clear selection
                        doc.selection = null;
                    } catch (e) {
                        // If grouping fails, try alternative method
                        try {
                            var colorGroup = doc.groupItems.add();
                            colorGroup.name = "Color Labels";
                            for (var g = 0; g < createdObjects.length; g++) {
                                createdObjects[g].moveToBeginning(colorGroup);
                            }
                        } catch (e2) {
                            // If both methods fail, continue without grouping
                        }
                    }
                }
                
                // Restore original selection
                doc.selection = originalSelection;
                
            } // End of colorNames.length > 0 check
        } // End of selection.length > 0 check
        
    } catch (error) {
        alert("Error creating color labels: " + error.toString());
    }
}