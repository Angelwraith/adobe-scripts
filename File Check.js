#target illustrator

/*@METADATA{
  "name": "File Check",
  "description": "Production file readiness analysis",
  "version": "2.5",
  "target": "illustrator",
  "tags": ["file", "check", "report"]
}@END_METADATA*/

// Main entry point with proper error checking
try {
    if (app.documents.length == 0) {
        alert("Please open at least one document first.");
    } else {
        showDocumentSelectionDialog();
    }
} catch (e) {
    alert("Startup error: " + e.toString());
}

function showDocumentSelectionDialog() {
    var dialog = new Window("dialog", "File Analysis - Select Documents");
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.spacing = 15;
    dialog.margins = 20;
    dialog.preferredSize.width = 500;
    dialog.preferredSize.height = 175;
    
    var titleText = dialog.add("statictext", undefined, "Select documents to analyze:");
    titleText.graphics.font = ScriptUI.newFont("dialog", "Bold", 14);
    
    dialog.add("panel");
    
    // Document selection - NO Analysis Options panel
    var documentsGroup = dialog.add("group");
    documentsGroup.orientation = "column";
    documentsGroup.alignChildren = "fill";
    
    var listPanel = documentsGroup.add("panel");
    listPanel.orientation = "column";
    listPanel.alignChildren = "fill";
    listPanel.margins = 10;
    listPanel.preferredSize.height = 180;
    
    var checkboxes = [];
    for (var i = 0; i < app.documents.length; i++) {
        var doc = app.documents[i];
        var checkboxGroup = listPanel.add("group");
        checkboxGroup.orientation = "row";
        checkboxGroup.alignChildren = "left";
        
        var checkbox = checkboxGroup.add("checkbox", undefined, doc.name);
        checkbox.value = true;
        checkbox.preferredSize.width = 400;
        
        checkboxes.push({
            checkbox: checkbox,
            document: doc
        });
    }
    
    dialog.add("panel");
    
    var selectionGroup = dialog.add("group");
    selectionGroup.alignment = "center";
    selectionGroup.spacing = 10;
    
    var selectAllButton = selectionGroup.add("button", undefined, "Select All");
    selectAllButton.preferredSize.width = 80;
    var selectCurrentButton = selectionGroup.add("button", undefined, "Select Current");
    selectCurrentButton.preferredSize.width = 90;
    var selectNoneButton = selectionGroup.add("button", undefined, "Select None");
    selectNoneButton.preferredSize.width = 80;
    
    selectAllButton.onClick = function() {
        for (var i = 0; i < checkboxes.length; i++) {
            checkboxes[i].checkbox.value = true;
        }
    };
    
    selectCurrentButton.onClick = function() {
        var activeDocName = app.activeDocument.name;
        for (var i = 0; i < checkboxes.length; i++) {
            checkboxes[i].checkbox.value = (checkboxes[i].document.name === activeDocName);
        }
    };
    
    selectNoneButton.onClick = function() {
        for (var i = 0; i < checkboxes.length; i++) {
            checkboxes[i].checkbox.value = false;
        }
    };
    
    dialog.add("panel");
    
    // THREE BUTTONS - No radio buttons anywhere
    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "center";
    buttonGroup.spacing = 10;
    
    var cancelButton = buttonGroup.add("button", undefined, "Cancel");
    var quickButton = buttonGroup.add("button", undefined, "Quick Analysis");
    var ppiButton = buttonGroup.add("button", undefined, "Full Analysis (with PPI)");
    
    var result = null;
    
    cancelButton.onClick = function() {
        result = null;
        dialog.close();
    };
    
    quickButton.onClick = function() {
        var selectedDocuments = [];
        for (var i = 0; i < checkboxes.length; i++) {
            if (checkboxes[i].checkbox.value) {
                selectedDocuments.push(checkboxes[i].document);
            }
        }
        
        if (selectedDocuments.length === 0) {
            alert("Please select at least one document to analyze.");
            return;
        }
        
        result = {
            documents: selectedDocuments,
            includePPI: false
        };
        dialog.close();
    };
    
    ppiButton.onClick = function() {
        var selectedDocuments = [];
        for (var i = 0; i < checkboxes.length; i++) {
            if (checkboxes[i].checkbox.value) {
                selectedDocuments.push(checkboxes[i].document);
            }
        }
        
        if (selectedDocuments.length === 0) {
            alert("Please select at least one document to analyze.");
            return;
        }
        
        result = {
            documents: selectedDocuments,
            includePPI: true
        };
        dialog.close();
    };
    
    dialog.show();
    
    if (result && result.documents && result.documents.length > 0) {
        runAnalysis(result.documents, result.includePPI);
    }
}
function runAnalysis(documents, includePPI) {
    try {
        var overallStartTime = new Date().getTime();
        var timingReport = [];
        var totalImageCount = 0;
        var allLowResImages = [];
        var allImageDetails = []; // NEW: Store ALL image details
        // Removed debug functionality
        var totalCutThroughPaths = 0;
        var allCutThroughSizes = {};
        var documentNames = [];
        var individualDocumentResults = [];
        
        // Get the scale factor for Large Canvas handling (like in PRIME script)
        var scaleFactor = 1; // Default to 1 if not defined
        try {
            // Get scaleFactor from the active document like PRIME script does
            scaleFactor = app.activeDocument.scaleFactor || 1;
        } catch (e) {
            scaleFactor = 1; // Fallback if scaleFactor doesn't exist
        }
        
        // Store original active document
        var originalActiveDoc = app.activeDocument;
        
        // Process each document
        for (var docIndex = 0; docIndex < documents.length; docIndex++) {
            var doc = documents[docIndex];
            documentNames.push(doc.name);
            
            var docStartTime = new Date().getTime();
            
            // Activate document for analysis
            var documentActivated = false;
            var targetDoc = null;
            
            for (var activateIndex = 0; activateIndex < app.documents.length; activateIndex++) {
                if (app.documents[activateIndex].name === doc.name) {
                    targetDoc = app.documents[activateIndex];
                    app.activeDocument = targetDoc;
                    documentActivated = true;
                    break;
                }
            }
            
            if (!documentActivated) {
                alert("ERROR: Could not activate document: " + doc.name);
                continue;
            }
            
            // Force redraw to ensure document is active
            app.redraw();
            $.sleep(100);
            
            var docResult = analyzeDocument(targetDoc, includePPI, scaleFactor);
            var analysisTime = new Date().getTime() - docStartTime;
            
			// Store individual document results
            var individualResult = {
                name: doc.name,
                analysisTime: analysisTime,
                rasterCount: docResult.rasterCount,
                lowResImages: docResult.lowResImages || [],
                allImageDetails: docResult.allImageDetails || [],
                cutThroughSizes: docResult.cutThroughSizes || {},
                totalCutThroughPaths: docResult.totalCutThroughPaths || 0,
                scaleFactor: scaleFactor
            };
            individualDocumentResults.push(individualResult);
            
            // Aggregate results
            totalImageCount += docResult.rasterCount;
            totalCutThroughPaths += docResult.totalCutThroughPaths;
            
            // Merge cut through sizes
            for (var size in docResult.cutThroughSizes) {
                if (allCutThroughSizes[size]) {
                    allCutThroughSizes[size] += docResult.cutThroughSizes[size];
                } else {
                    allCutThroughSizes[size] = docResult.cutThroughSizes[size];
                }
            }
            
            // Merge low res images
            if (docResult.lowResImages) {
                for (var i = 0; i < docResult.lowResImages.length; i++) {
                    allLowResImages.push({
                        document: doc.name,
                        text: docResult.lowResImages[i].text
                    });
                }
            }
            
            // NEW: Merge all image details
            if (docResult.allImageDetails) {
                for (var i = 0; i < docResult.allImageDetails.length; i++) {
                    allImageDetails.push({
                        document: doc.name,
                        text: docResult.allImageDetails[i].text
                    });
                }
            }
            
            timingReport.push("Document '" + doc.name + "': " + analysisTime + "ms");
            
            // Cleanup between documents
            if (docIndex < documents.length - 1) {
                try {
                    for (var cleanupIndex = 0; cleanupIndex < app.documents.length; cleanupIndex++) {
                        try {
                            app.documents[cleanupIndex].selection = null;
                        } catch (e) {}
                    }
                    app.redraw();
                    $.gc();
                    $.sleep(200);
                } catch (cleanupError) {}
            }
        }
        
        // Restore original active document
        try {
            app.activeDocument = originalActiveDoc;
        } catch (e) {}
        
        var totalTime = new Date().getTime() - overallStartTime;
        
// Build report text
        var reportText = buildReport(documents.length, documentNames, totalTime, timingReport, 
                                   totalImageCount, allLowResImages, allImageDetails, allCutThroughSizes, 
                                   totalCutThroughPaths, includePPI, individualDocumentResults);
        
        // Store analysis data
        var analysisData = {
            reportText: reportText,
            individualResults: individualDocumentResults,
            totalDocuments: documents.length,
            totalPaths: totalCutThroughPaths,
            totalTime: Math.round(totalTime / 1000),
            hasLowRes: includePPI && allLowResImages.length > 0,
            allCutThroughSizes: allCutThroughSizes,
            allImageDetails: allImageDetails, // NEW: Pass all image details
        };
        
        // Show results
        showResultsDialog(analysisData);
        
    } catch (error) {
        alert("Error in analysis: " + error.toString());
    }
}

function analyzeDocument(doc, includePPI, scaleFactor) {
    function findPathsWithCutThroughColor(doc) {
	var originalSelection = doc.selection;
        
        var cutThroughSizes = {};
        var totalPaths = 0;
        
        try {
            var targetSpot = null;
            for (var i = 0; i < doc.spots.length; i++) {
                if (doc.spots[i].name == "CutThrough2-Outside") {
                    targetSpot = doc.spots[i];
                    break;
                }
            }
            
            if (!targetSpot) {
                return {
                    cutThroughSizes: cutThroughSizes,
                    totalCutThroughPaths: 0
                };
            }
            
            var spotColor = new SpotColor();
            spotColor.spot = targetSpot;
            
            doc.selection = null;
            doc.defaultFillColor = spotColor;
            app.executeMenuCommand("Find Fill Color menu item");
            
            var fillSelection = [];
            for (var i = 0; i < doc.selection.length; i++) {
                fillSelection.push(doc.selection[i]);
            }
            
            doc.selection = null;
            doc.defaultStrokeColor = spotColor;
            app.executeMenuCommand("Find Stroke Color menu item");
            
            var strokeSelection = [];
            for (var i = 0; i < doc.selection.length; i++) {
                strokeSelection.push(doc.selection[i]);
            }
            
            var allPaths = [];
            var processedItems = [];
            
            for (var i = 0; i < fillSelection.length; i++) {
                if (fillSelection[i].typename == "PathItem") {
                    allPaths.push(fillSelection[i]);
                    processedItems.push(fillSelection[i]);
                }
            }
            
            for (var i = 0; i < strokeSelection.length; i++) {
                if (strokeSelection[i].typename == "PathItem") {
                    var alreadyAdded = false;
                    for (var j = 0; j < processedItems.length; j++) {
                        if (processedItems[j] === strokeSelection[i]) {
                            alreadyAdded = true;
                            break;
                        }
                    }
                    if (!alreadyAdded) {
                        allPaths.push(strokeSelection[i]);
                    }
                }
            }
            
            for (var i = 0; i < allPaths.length; i++) {
                var size = getPathDimensions(allPaths[i]);
                if (cutThroughSizes[size]) {
                    cutThroughSizes[size]++;
                } else {
                    cutThroughSizes[size] = 1;
                }
                totalPaths++;
            }
            
        } catch (e) {
        } finally {
			try {
                doc.selection = originalSelection;
            } catch (e) {}
        }
        
        return {
            cutThroughSizes: cutThroughSizes,
            totalCutThroughPaths: totalPaths
        };
    }

    function pointsToInches(points, doc) {
        try {
            var scaleFactor = doc.scaleFactor;
            if (!scaleFactor || scaleFactor <= 0 || isNaN(scaleFactor)) {
                scaleFactor = 1;
            }
            return Math.round(((points / 72) * scaleFactor) * 100) / 100;
        } catch (e) {
            return Math.round((points / 72) * 100) / 100;
        }
    }

    function getPathDimensions(pathItem) {
        var bounds = pathItem.geometricBounds;
        var width = pointsToInches(bounds[2] - bounds[0]);
        var height = pointsToInches(bounds[1] - bounds[3]);
        
        // Round to nearest 1/8" (0.125)
        width = Math.round(width * 8) / 8;
        height = Math.round(height * 8) / 8;
        
        if (width <= height) {
            return width + '"x' + height + '"';
        } else {
            return height + '"x' + width + '"';
        }
    }

    function hasCutThroughColor(item) {
        try {
            if (item.filled && item.fillColor) {
                if (item.fillColor.typename == "SpotColor" &&
                     item.fillColor.spot &&
                     item.fillColor.spot.name == "CutThrough2-Outside") {
                    return true;
                }
            }
            if (item.stroked && item.strokeColor) {
                if (item.strokeColor.typename == "SpotColor" &&
                     item.strokeColor.spot &&
                     item.strokeColor.spot.name == "CutThrough2-Outside") {
                    return true;
                }
            }
        } catch (e) {}
        return false;
    }

    function processPageItems(items, cutThroughSizes, processedItems) {
        if (!processedItems) {
            processedItems = [];
        }
        
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            
            var itemFound = false;
            for (var j = 0; j < processedItems.length; j++) {
                if (processedItems[j] === item) {
                    itemFound = true;
                    break;
                }
            }
            if (itemFound) continue;
            
            if (item.typename == "PathItem" && hasCutThroughColor(item)) {
                var size = getPathDimensions(item);
                if (cutThroughSizes[size]) {
                    cutThroughSizes[size]++;
                } else {
                    cutThroughSizes[size] = 1;
                }
                processedItems.push(item);
            }
            
            if (item.typename == "GroupItem" && item.pageItems.length > 0) {
                processPageItems(item.pageItems, cutThroughSizes, processedItems);
            }
        }
        
        return processedItems;
    }

// Image analysis - FIXED to show ALL images and better PPI detection
    var rasterCount = 0;
    var lowResImages = [];
    var allImageDetails = []; // NEW: Store details for ALL images
    
    // FAST PRE-CHECK: Skip image analysis if document has no images
    var hasImages = false;
    try {
        if (doc.rasterItems.length > 0 || doc.placedItems.length > 0) {
            hasImages = true;
        }
    } catch (e) {
        hasImages = true;
    }
    
    if (includePPI && hasImages) {
        try {
            var processedItems = [];

            function processImageItem(item) {
                if (item.typename == "RasterItem" || item.typename == "PlacedItem") {
                    var alreadyProcessed = false;
                    for (var k = 0; k < processedItems.length; k++) {
                        if (processedItems[k] === item) {
                            alreadyProcessed = true;
                            break;
                        }
                    }
                    
                    if (!alreadyProcessed) {
                        processedItems.push(item);
                        rasterCount++;
                        
                        try {
                            // Get basic image info with Large Canvas correction
                            var bounds = item.geometricBounds;
                            // Apply scale factor correction for Large Canvas mode  
                            var actualScaleFactor = (scaleFactor && scaleFactor > 0) ? scaleFactor : 1;
                            var docWidthInches = Math.round(((bounds[2] - bounds[0]) / 72 * scaleFactor) * 100) / 100;
                            var docHeightInches = Math.round(((bounds[1] - bounds[3]) / 72 * scaleFactor) * 100) / 100;
                            
                            var imageInfo = {
                                type: item.typename,
                                bounds: bounds,
                                docWidth: docWidthInches,
                                docHeight: docHeightInches
                            };
                            
                            // Try multiple methods to get PPI/resolution info
                            var estimatedPPI = "Unknown";
                            var resolutionMethod = "No method available";
                            
                            // Method 1: Try to get actual resolution for RasterItems
                            if (item.typename == "RasterItem") {
                                try {
                                    // Check if it's linked or embedded
                                    var fileStatus = "Embedded";
                                    try {
                                        if (item.file && item.file.exists) {
                                            fileStatus = "Linked: " + item.file.name;
                                        }
                                    } catch (fileError) {
                                        // File property doesn't exist or accessible - definitely embedded
                                        fileStatus = "Embedded";
                                    }
                                    
                                    resolutionMethod = fileStatus;
                                    
                                    // Calculate from transformation matrix (works for both linked and embedded)
                                    try {
                                        var matrix = item.matrix;
                                        if (matrix && matrix.mValueA !== undefined && matrix.mValueD !== undefined) {
                                            var scaleX = Math.abs(matrix.mValueA);
                                            var scaleY = Math.abs(matrix.mValueD);
                                            var avgScale = (scaleX + scaleY) / 2;
                                            if (avgScale > 0) {
                                                // Apply Large Canvas correction to PPI calculation
                                                var basePPI = 72 / avgScale;
                                                // In Large Canvas mode, images appear smaller but are actually higher resolution
                                                // scaleFactor is typically 0.1 for Large Canvas, meaning 10x correction needed
                                                var correctedPPI = Math.round(basePPI / scaleFactor);
                                                estimatedPPI = correctedPPI;
                                                resolutionMethod += " (Matrix: scaleX=" + scaleX.toFixed(3) + ", scaleY=" + scaleY.toFixed(3) + ", avg=" + avgScale.toFixed(3);
                                                if (actualScaleFactor !== 1) {
                                                    resolutionMethod += ", Large Canvas factor=" + actualScaleFactor + ", corrected PPI=" + correctedPPI;
                                                }
                                                resolutionMethod += ")";
                                            } else {
                                                resolutionMethod += " (Matrix: invalid scale values)";
                                            }
                                        } else {
                                            resolutionMethod += " (Matrix: no matrix data)";
                                        }
                                    } catch (matrixError) {
                                        resolutionMethod += " (Matrix error: " + matrixError.toString() + ")";
                                    }
                                } catch (rasterError) {
                                    resolutionMethod = "RasterItem analysis error: " + rasterError.toString();
                                }
                            }
                            
                            // Method 2: Try for PlacedItems (linked images)
                            else if (item.typename == "PlacedItem") {
                                try {
                                    var fileStatus = "Unknown";
                                    try {
                                        if (item.file && item.file.exists) {
                                            fileStatus = "Linked: " + item.file.name;
                                        } else {
                                            fileStatus = "Missing link";
                                        }
                                    } catch (fileError) {
                                        fileStatus = "No file reference";
                                    }
                                    
                                    resolutionMethod = fileStatus;
                                    
                                    // Calculate from transformation matrix
                                    try {
                                        var matrix = item.matrix;
                                        if (matrix && matrix.mValueA !== undefined && matrix.mValueD !== undefined) {
                                            var scaleX = Math.abs(matrix.mValueA);
                                            var scaleY = Math.abs(matrix.mValueD);
                                            var avgScale = (scaleX + scaleY) / 2;
                                            if (avgScale > 0) {
                                                estimatedPPI = Math.round(72 / avgScale);
                                                resolutionMethod += " (Matrix: scaleX=" + scaleX.toFixed(3) + ", scaleY=" + scaleY.toFixed(3) + ", avg=" + avgScale.toFixed(3) + ")";
                                            } else {
                                                resolutionMethod += " (Matrix: invalid scale values)";
                                            }
                                        } else {
                                            resolutionMethod += " (Matrix: no matrix data)";
                                        }
                                    } catch (matrixError) {
                                        resolutionMethod += " (Matrix error: " + matrixError.toString() + ")";
                                    }
                                } catch (placedError) {
                                    resolutionMethod = "PlacedItem analysis error: " + placedError.toString();
                                }
                            }
                            
                            // Build comprehensive image text with RAW and CORRECTED values for debugging
                            var imageText = 'Img#' + rasterCount + ' (' + item.typename + '): SIZE: ' + docWidthInches + '"x' + docHeightInches + '"';
                            
                            if (estimatedPPI !== "Unknown") {
                                imageText += ', PPI: ~' + estimatedPPI;
                            } else {
                                imageText += ', PPI: Unknown';
                            }
                            
                            imageText += ' [' + resolutionMethod + ']';
                            
                            // Add to ALL images list
                            allImageDetails.push({text: imageText});
                            
                            // Check if low resolution (only if we got a numeric PPI)
                            var isLowRes = false;
                            if (estimatedPPI !== "Unknown" && typeof estimatedPPI === "number") {
                                isLowRes = (estimatedPPI < 72);
                                if (isLowRes) {
                                    imageText += ' LOW RESOLUTION!';
									lowResImages.push({text: imageText});
                                }
                            }
                            
                        } catch (imageError) {
                            // If we can't analyze the image, still record it
                            var errorText = 'Image ' + rasterCount + ' (' + item.typename + '): Analysis error - ' + imageError.toString();
                            allImageDetails.push({text: errorText});
                        }
                    }
                }
                
                if (item.typename == "GroupItem" && item.pageItems.length > 0) {
                    for (var j = 0; j < item.pageItems.length; j++) {
                        processImageItem(item.pageItems[j]);
                    }
                }
            }

            for (var i = 0; i < doc.pageItems.length; i++) {
                processImageItem(doc.pageItems[i]);
            }
        } catch (e) {
            // If PPI analysis fails completely, still record that we tried
            allImageDetails.push({text: "PPI Analysis failed: " + e.toString()});
        }
} else if (hasImages) {
        // Quick analysis - just count images using built-in collections (INSTANT)
        try {
            rasterCount = doc.rasterItems.length + doc.placedItems.length;
            if (rasterCount > 0) {
                allImageDetails.push({text: "Quick scan found " + rasterCount + " images (no details in quick mode)"});
            }
        } catch (e) {
            allImageDetails.push({text: "Image counting failed: " + e.toString()});
        }
    }

// CutThrough analysis - FAST METHOD using Find Color
    var cutThroughResults = findPathsWithCutThroughColor(doc);
    var cutThroughSizes = cutThroughResults.cutThroughSizes;
    var totalCutThroughPaths = cutThroughResults.totalCutThroughPaths;


    return {
        rasterCount: rasterCount,
        lowResImages: lowResImages,
        allImageDetails: allImageDetails, // NEW: Return all image details
        cutThroughSizes: cutThroughSizes,
        totalCutThroughPaths: totalCutThroughPaths,
    };
}

function showResultsDialog(analysisData) {
    var dialog = new Window("dialog", "Analysis Results");
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.spacing = 10;
    dialog.margins = 15;
    dialog.preferredSize.width = 700; // Increased width for more detail
	dialog.layout.layout(true);
	dialog.layout.resize();
	dialog.preferredSize.height = 1200;
	dialog.maximumSize.height = 1200;
	dialog.minimumSize.height = 1200;
    
    // Summary header
    var headerPanel = dialog.add("panel", undefined, "Analysis Summary");
    headerPanel.orientation = "column";
    headerPanel.alignChildren = "fill";
    headerPanel.margins = 10;
    
    var summaryGroup = headerPanel.add("group");
    summaryGroup.orientation = "row";
    summaryGroup.alignment = "fill";
    
    var docsLabel = summaryGroup.add("statictext", undefined, "Documents: " + analysisData.totalDocuments);
    var pathsLabel = summaryGroup.add("statictext", undefined, "Cut Paths: " + analysisData.totalPaths);
    var timeLabel = summaryGroup.add("statictext", undefined, "Time: " + analysisData.totalTime + "s");
    
    // Results tabs - increased height to fill more space
	var tabGroup = dialog.add("tabbedpanel");
	tabGroup.alignChildren = "fill";
	tabGroup.preferredSize.height = 1075;
	tabGroup.maximumSize.height = 1075;
    
    // Summary tab
    var summaryTab = tabGroup.add("tab", undefined, "Summary");
    summaryTab.orientation = "column";
    summaryTab.alignChildren = "fill";
    
    var summaryText = summaryTab.add("edittext", undefined, "", {multiline: true, scrolling: true, readonly: true});
    summaryText.alignment = ["fill", "fill"];
    
    // Build summary content (concise) with low-res warning if needed
    var summaryContent = [];
    
// Add low-res warning at top if needed  
if (analysisData.hasLowRes) {
    // Calculate low res count on the fly
    var summaryLowResCount = 0;
    if (analysisData.individualResults) {
        for (var lr = 0; lr < analysisData.individualResults.length; lr++) {
            if (analysisData.individualResults[lr].lowResImages) {
                summaryLowResCount += analysisData.individualResults[lr].lowResImages.length;
            }
        }
    }
    
    summaryContent.push("ATTENTION: " + summaryLowResCount + " LOW RESOLUTION IMAGES FOUND");
    summaryContent.push("Images below 72 PPI may not print clearly!");
    summaryContent.push("CHECK INDIVIDUAL DOCUMENT TABS FOR DETAILS");
    summaryContent.push("");
}
    
    summaryContent.push("ANALYSIS SUMMARY");
    summaryContent.push("Documents: " + analysisData.totalDocuments);
    summaryContent.push("Total Cut Paths: " + analysisData.totalPaths);
    summaryContent.push("");
    summaryContent.push("CUTTHROUGH BREAKDOWN");
    
    // Sort and format sizes with proper tab alignment
    var sizes = [];
    for (var size in analysisData.allCutThroughSizes) {
        sizes.push(size);
    }
    sizes.sort();
    
    for (var s = 0; s < sizes.length; s++) {
        var size = sizes[s];
        var count = analysisData.allCutThroughSizes[size];
        var countStr = count.toString() + "/ea";
        // Use consistent spacing like "14/ea" line - pad to 6 characters then tab
        var paddedCount = (countStr + "      ").substr(0, 6);
        summaryContent.push(paddedCount + "\t" + size);
    }
    
    summaryText.text = summaryContent.join("\n");
    
    // Add detailed summary tab
    var detailedTab = tabGroup.add("tab", undefined, "Details");
    detailedTab.orientation = "column";
    detailedTab.alignChildren = "fill";
    
    var detailedText = detailedTab.add("edittext", undefined, "", {multiline: true, scrolling: true, readonly: true});
    detailedText.alignment = ["fill", "fill"];
    
// Add prominent low-res warning header if needed
var detailedContent = analysisData.reportText;
if (analysisData.hasLowRes) {
    // Calculate low res count on the fly
    var detailedLowResCount = 0;
    if (analysisData.individualResults) {
        for (var lr = 0; lr < analysisData.individualResults.length; lr++) {
            if (analysisData.individualResults[lr].lowResImages) {
                detailedLowResCount += analysisData.individualResults[lr].lowResImages.length;
            }
        }
    }
    
    var warningHeader = "ATTENTION: " + detailedLowResCount + " LOW RESOLUTION IMAGES FOUND\n";
    warningHeader += "Images below 72 PPI may not print clearly!\n";
    warningHeader += "CHECK INDIVIDUAL DOCUMENT TABS FOR DETAILS\n\n";
    detailedContent = warningHeader + detailedContent;
}
    
    detailedText.text = detailedContent;
    
// Individual document tabs
    if (analysisData.individualResults) {
        for (var d = 0; d < analysisData.individualResults.length; d++) { // Show all documents
            var docResult = analysisData.individualResults[d];
            var docTab = tabGroup.add("tab", undefined, "Doc " + (d + 1));
            docTab.orientation = "column";
            docTab.alignChildren = "fill";
            
            var docText = docTab.add("edittext", undefined, "", {multiline: true, scrolling: true, readonly: true});
            docText.alignment = ["fill", "fill"];
            
            var docReport = [];
            docReport.push(docResult.name);
            docReport.push("Processing: " + docResult.analysisTime + "ms");
            docReport.push("Images: " + docResult.rasterCount);
            docReport.push("Cut Paths: " + docResult.totalCutThroughPaths);
            docReport.push("");
            
            if (docResult.totalCutThroughPaths > 0) {
                docReport.push("CUTTHROUGH BREAKDOWN");
                var docSizes = [];
                for (var size in docResult.cutThroughSizes) {
                    docSizes.push(size);
                }
                docSizes.sort();
                
                for (var s = 0; s < docSizes.length; s++) {
                    var size = docSizes[s];
                    var count = docResult.cutThroughSizes[size];
                    var countStr = count.toString() + "/ea";
                    // Use consistent spacing - pad to 6 characters then tab
                    var paddedCount = (countStr + "      ").substr(0, 6);
                    docReport.push(paddedCount + "\t" + size);
                }
            } else {
                docReport.push("No CutThrough paths found");
            }
            
            // NEW: Show ALL image details in individual document tabs
            if (docResult.allImageDetails && docResult.allImageDetails.length > 0) {
                docReport.push("");
                docReport.push("ALL IMAGE DETAILS");
                for (var img = 0; img < docResult.allImageDetails.length; img++) {
                    docReport.push(docResult.allImageDetails[img].text);
                }
            }
            
            // Add low-res image details if available (keep separate for emphasis)
            if (docResult.lowResImages && docResult.lowResImages.length > 0) {
                docReport.push("");
                docReport.push("*** LOW RESOLUTION IMAGES ***");
                for (var img = 0; img < docResult.lowResImages.length; img++) {
                    docReport.push(docResult.lowResImages[img].text);
                }
            }
            
            docText.text = docReport.join("\n");
        }
    }
    
	// Buttons - force to bottom
	var buttonGroup = dialog.add("group");
	buttonGroup.maximumSize.height = 50;
    buttonGroup.alignment = "center";
    buttonGroup.spacing = 10;
    
    var exportBtn = buttonGroup.add("button", undefined, "Export Report");
    var closeBtn = buttonGroup.add("button", undefined, "Close");
    
    exportBtn.onClick = function() {
        try {
            var saveFile = File.saveDialog("Save report as text file", "Text files:*.txt");
            if (saveFile) {
                saveFile.open("w");
                saveFile.write(analysisData.reportText);
                saveFile.close();
                alert("Report saved successfully!");
            }
        } catch (e) {
            alert("Error saving file: " + e.toString());
        }
    };
    
    closeBtn.onClick = function() {
        dialog.close();
    };
    
    dialog.show();
}

function buildReport(docCount, documentNames, totalTime, timingReport, 
                    totalImageCount, allLowResImages, allImageDetails, allCutThroughSizes, 
                    totalCutThroughPaths, includePPI, individualDocumentResults) {
    var report = [];
    report.push("=== MULTI-DOCUMENT ANALYSIS SUMMARY ===");
    report.push("");
    
    // Check if any documents are Large Canvas
    var hasLargeCanvas = false;
    for (var i = 0; i < individualDocumentResults.length; i++) {
        if (individualDocumentResults[i].scaleFactor && individualDocumentResults[i].scaleFactor !== 1) {
            hasLargeCanvas = true;
            break;
        }
    }
    
    if (hasLargeCanvas) {
        report.push("*** LARGE CANVAS DOCUMENTS DETECTED ***");
        report.push("Some documents use Large Canvas mode. Measurements have been corrected.");
        report.push("");
    }
    
    report.push("Analysis Type: " + (includePPI ? "Full Analysis (with PPI calculations)" : "Quick Analysis (no PPI calculations)"));
    report.push("Documents Analyzed: " + docCount);
    report.push("Total Processing Time: " + Math.round(totalTime / 1000) + " seconds");
    report.push("");
    
    report.push("Documents Processed:");
    for (var i = 0; i < documentNames.length; i++) {
        report.push("  " + (i + 1) + ". " + documentNames[i]);
    }
    report.push("");
    
    report.push("=== KEY METRICS ===");
    report.push("Total Images: " + totalImageCount);
    if (includePPI && allLowResImages.length > 0) {
        report.push("Low Resolution Images (< 72 PPI): " + allLowResImages.length + " ** WARNING **");
    } else if (includePPI) {
        report.push("Low Resolution Images (< 72 PPI): 0 (All images meet requirements)");
    }
report.push("Total CutThrough2-Outside Paths: " + totalCutThroughPaths);
    report.push("");
    
// Per-document breakdown
    report.push("=== PER-DOCUMENT BREAKDOWN ===");
    
// Sort documents using localeCompare with numeric option
    var sortedResults = [];
    for (var i = 0; i < individualDocumentResults.length; i++) {
        sortedResults.push(individualDocumentResults[i]);
    }
    sortedResults.sort(function(a, b) {
        return a.name.localeCompare(b.name, undefined, {numeric: true});
    });
    
    for (var i = 0; i < sortedResults.length; i++) {
        var docResult = sortedResults[i];
        var docResult = individualDocumentResults[i];
        report.push("");
        report.push("Document: " + docResult.name);
        if (docResult.scaleFactor && docResult.scaleFactor !== 1) {
            report.push("  ** Large Canvas (Scale Factor: " + docResult.scaleFactor + ") **");
        }       
        if (docResult.totalCutThroughPaths > 0) {
            var docSizes = [];
            for (var size in docResult.cutThroughSizes) {
                docSizes.push(size);
            }
            docSizes.sort();
            for (var s = 0; s < docSizes.length; s++) {
                var size = docSizes[s];
                var count = docResult.cutThroughSizes[size];
                report.push("    " + count + " x " + size);
            }
        } else {
            report.push("  No CutThrough paths found");
        }
    }
    report.push("");
    
    // NEW: Add all image details to report
    if (includePPI && allImageDetails && allImageDetails.length > 0) {
        report.push("=== ALL IMAGE DETAILS ===");
        for (var i = 0; i < allImageDetails.length; i++) {
            report.push(allImageDetails[i].text);
        }
        report.push("");
    }
    
    // CutThrough breakdown
    if (totalCutThroughPaths > 0) {
        report.push("=== CUTTHROUGH2-OUTSIDE BREAKDOWN ===");
        var sortedSizes = [];
        for (var size in allCutThroughSizes) {
            sortedSizes.push(size);
        }
        sortedSizes.sort();
        
        for (var i = 0; i < sortedSizes.length; i++) {
            var size = sortedSizes[i];
            var count = allCutThroughSizes[size];
            report.push(count + "\t" + size);
        }
        report.push("");
    }
    
    // Performance timing
    report.push("=== PERFORMANCE TIMING ===");
    report.push("Overall Processing: " + Math.round(totalTime) + "ms (" + Math.round(totalTime / 1000) + "s)");
    report.push("Average per Document: " + Math.round(totalTime / docCount) + "ms");
    report.push("");
    report.push("Individual Document Times:");
    for (var i = 0; i < timingReport.length; i++) {
        report.push("  " + timingReport[i]);
    }
    
    return report.join("\n");
}