#target illustrator

/*@METADATA{
  "name": "File Check",
  "description": "Production file readiness analysis",
  "version": "2.6",
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
        var allImageDetails = [];
        var totalCutThroughPaths = 0;
        var totalCompoundPaths = 0;
        var allCutThroughSizes = {};
        var allCompoundPathWarnings = [];
        var documentNames = [];
        var individualDocumentResults = [];
        
        // Get the scale factor for Large Canvas handling
        var scaleFactor = 1;
        try {
            scaleFactor = app.activeDocument.scaleFactor || 1;
        } catch (e) {
            scaleFactor = 1;
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
                compoundPathWarnings: docResult.compoundPathWarnings || [],
                totalCompoundPaths: docResult.totalCompoundPaths || 0,
                scaleFactor: scaleFactor
            };
            individualDocumentResults.push(individualResult);
            
            // Aggregate results
            totalImageCount += docResult.rasterCount;
            totalCutThroughPaths += docResult.totalCutThroughPaths;
            totalCompoundPaths += docResult.totalCompoundPaths;
            
            // Merge cut through sizes
            for (var size in docResult.cutThroughSizes) {
                if (allCutThroughSizes[size]) {
                    allCutThroughSizes[size] += docResult.cutThroughSizes[size];
                } else {
                    allCutThroughSizes[size] = docResult.cutThroughSizes[size];
                }
            }
            
            // Merge compound path warnings
            if (docResult.compoundPathWarnings) {
                for (var i = 0; i < docResult.compoundPathWarnings.length; i++) {
                    allCompoundPathWarnings.push({
                        document: doc.name,
                        text: docResult.compoundPathWarnings[i].text
                    });
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
            
            // Merge all image details
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
                                   totalCutThroughPaths, totalCompoundPaths, allCompoundPathWarnings,
                                   includePPI, individualDocumentResults);
        
        // Store analysis data
        var analysisData = {
            reportText: reportText,
            individualResults: individualDocumentResults,
            totalDocuments: documents.length,
            totalPaths: totalCutThroughPaths,
            totalCompoundPaths: totalCompoundPaths,
            totalTime: Math.round(totalTime / 1000),
            hasLowRes: includePPI && allLowResImages.length > 0,
            hasCompoundPaths: totalCompoundPaths > 0,
            allCutThroughSizes: allCutThroughSizes,
            allImageDetails: allImageDetails,
            allCompoundPathWarnings: allCompoundPathWarnings
        };
        
        // Show results
        showResultsDialog(analysisData);
        
    } catch (error) {
        alert("Error in analysis: " + error.toString());
    }
}

function analyzeDocument(doc, includePPI, scaleFactor) {
    function pointsToInches(points) {
        try {
            var docScaleFactor = doc.scaleFactor;
            if (!docScaleFactor || docScaleFactor <= 0 || isNaN(docScaleFactor)) {
                docScaleFactor = 1;
            }
            return Math.round(((points / 72) * docScaleFactor) * 100) / 100;
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

    function findPathsWithCutThroughColor(doc) {
        var originalSelection = doc.selection;
        
        var cutThroughSizes = {};
        var totalPaths = 0;
        var compoundPathWarnings = [];
        var totalCompoundPaths = 0;
        
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
                    totalCutThroughPaths: 0,
                    compoundPathWarnings: compoundPathWarnings,
                    totalCompoundPaths: 0
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
                } else if (fillSelection[i].typename == "CompoundPathItem") {
                    // COMPOUND PATH DETECTED - This is a production error!
                    totalCompoundPaths++;
                    var size = getPathDimensions(fillSelection[i]);
                    compoundPathWarnings.push({
                        text: "WARNING: CutThrough color on COMPOUND PATH (size: " + size + ") - This will cause production errors!"
                    });
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
                } else if (strokeSelection[i].typename == "CompoundPathItem") {
                    // Check if already counted
                    var alreadyCounted = false;
                    for (var j = 0; j < fillSelection.length; j++) {
                        if (fillSelection[j] === strokeSelection[i]) {
                            alreadyCounted = true;
                            break;
                        }
                    }
                    if (!alreadyCounted) {
                        totalCompoundPaths++;
                        var size = getPathDimensions(strokeSelection[i]);
                        compoundPathWarnings.push({
                            text: "WARNING: CutThrough color on COMPOUND PATH (size: " + size + ") - This will cause production errors!"
                        });
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
            totalCutThroughPaths: totalPaths,
            compoundPathWarnings: compoundPathWarnings,
            totalCompoundPaths: totalCompoundPaths
        };
    }

    // Image analysis
    var rasterCount = 0;
    var lowResImages = [];
    var allImageDetails = [];
    
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
                                    var fileStatus = "Embedded";
                                    try {
                                        if (item.file && item.file.exists) {
                                            fileStatus = "Linked: " + item.file.name;
                                        }
                                    } catch (fileError) {
                                        fileStatus = "Embedded";
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
                                                var basePPI = 72 / avgScale;
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
                            
                            // Method 2: Try for PlacedItems
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
                            
                            // Build comprehensive image text
                            var imageText = 'Img#' + rasterCount + ' (' + item.typename + '): SIZE: ' + docWidthInches + '"x' + docHeightInches + '"';
                            
                            if (estimatedPPI !== "Unknown") {
                                imageText += ', PPI: ~' + estimatedPPI;
                            } else {
                                imageText += ', PPI: Unknown';
                            }
                            
                            imageText += ' [' + resolutionMethod + ']';
                            
                            // Add to ALL images list
                            allImageDetails.push({text: imageText});
                            
                            // Check if low resolution
                            var isLowRes = false;
                            if (estimatedPPI !== "Unknown" && typeof estimatedPPI === "number") {
                                isLowRes = (estimatedPPI < 72);
                                if (isLowRes) {
                                    imageText += ' LOW RESOLUTION!';
                                    lowResImages.push({text: imageText});
                                }
                            }
                            
                        } catch (imageError) {
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
            allImageDetails.push({text: "PPI Analysis failed: " + e.toString()});
        }
    } else if (hasImages) {
        // Quick analysis - just count images
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
    var compoundPathWarnings = cutThroughResults.compoundPathWarnings;
    var totalCompoundPaths = cutThroughResults.totalCompoundPaths;

    return {
        rasterCount: rasterCount,
        lowResImages: lowResImages,
        allImageDetails: allImageDetails,
        cutThroughSizes: cutThroughSizes,
        totalCutThroughPaths: totalCutThroughPaths,
        compoundPathWarnings: compoundPathWarnings,
        totalCompoundPaths: totalCompoundPaths
    };
}

function showResultsDialog(analysisData) {
    var dialog = new Window("dialog", "Analysis Results");
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.spacing = 10;
    dialog.margins = 15;
    dialog.preferredSize.width = 700;
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
    
    // Add critical warnings if detected
    if (analysisData.hasCompoundPaths) {
        var warningGroup = headerPanel.add("group");
        warningGroup.orientation = "row";
        warningGroup.alignment = "fill";
        var compoundWarning = warningGroup.add("statictext", undefined, "COMPOUND PATHS: " + analysisData.totalCompoundPaths + " - PRODUCTION ERROR!");
        compoundWarning.graphics.foregroundColor = compoundWarning.graphics.newPen(compoundWarning.graphics.PenType.SOLID_COLOR, [1, 0, 0], 1);
        compoundWarning.graphics.font = ScriptUI.newFont("dialog", "Bold", 12);
    }
    
    if (analysisData.hasLowRes) {
        var lowResCount = 0;
        if (analysisData.individualResults) {
            for (var lr = 0; lr < analysisData.individualResults.length; lr++) {
                if (analysisData.individualResults[lr].lowResImages) {
                    lowResCount += analysisData.individualResults[lr].lowResImages.length;
                }
            }
        }
        var lowResGroup = headerPanel.add("group");
        lowResGroup.orientation = "row";
        lowResGroup.alignment = "fill";
        var lowResWarning = lowResGroup.add("statictext", undefined, "LOW RESOLUTION: " + lowResCount + " images below 72 PPI!");
        lowResWarning.graphics.foregroundColor = lowResWarning.graphics.newPen(lowResWarning.graphics.PenType.SOLID_COLOR, [1, 0, 0], 1);
        lowResWarning.graphics.font = ScriptUI.newFont("dialog", "Bold", 12);
    }
    
    // Results tabs
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
    
    // Build summary content with warnings
    var summaryContent = [];
    
    // Add compound path warning at top if needed
    if (analysisData.hasCompoundPaths) {
        summaryContent.push("*** CRITICAL PRODUCTION ERROR ***");
        summaryContent.push("COMPOUND PATHS WITH CUTTHROUGH COLOR DETECTED!");
        summaryContent.push("Found " + analysisData.totalCompoundPaths + " compound path(s) with CutThrough color");
        summaryContent.push("This WILL cause cutting machine failures!");
        summaryContent.push("MUST BE FIXED before production!");
        summaryContent.push("");
    }
    
    // Add low-res warning at top if needed  
    if (analysisData.hasLowRes) {
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
    
    // Sort and format sizes
    var sizes = [];
    for (var size in analysisData.allCutThroughSizes) {
        sizes.push(size);
    }
    sizes.sort();
    
    for (var s = 0; s < sizes.length; s++) {
        var size = sizes[s];
        var count = analysisData.allCutThroughSizes[size];
        var countStr = count.toString() + "/ea";
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
    
    // Add warnings to detailed view
    var detailedContent = analysisData.reportText;
    if (analysisData.hasCompoundPaths || analysisData.hasLowRes) {
        var warningHeader = "";
        
        if (analysisData.hasCompoundPaths) {
            warningHeader += "*** CRITICAL PRODUCTION ERROR ***\n";
            warningHeader += "COMPOUND PATHS WITH CUTTHROUGH COLOR DETECTED!\n";
            warningHeader += "Found " + analysisData.totalCompoundPaths + " compound path(s) with CutThrough color\n";
            warningHeader += "This WILL cause cutting machine failures!\n";
            warningHeader += "MUST BE FIXED before production!\n\n";
        }
        
        if (analysisData.hasLowRes) {
            var detailedLowResCount = 0;
            if (analysisData.individualResults) {
                for (var lr = 0; lr < analysisData.individualResults.length; lr++) {
                    if (analysisData.individualResults[lr].lowResImages) {
                        detailedLowResCount += analysisData.individualResults[lr].lowResImages.length;
                    }
                }
            }
            
            warningHeader += "ATTENTION: " + detailedLowResCount + " LOW RESOLUTION IMAGES FOUND\n";
            warningHeader += "Images below 72 PPI may not print clearly!\n";
            warningHeader += "CHECK INDIVIDUAL DOCUMENT TABS FOR DETAILS\n\n";
        }
        
        detailedContent = warningHeader + detailedContent;
    }
    
    detailedText.text = detailedContent;
    
    // Individual document tabs
    if (analysisData.individualResults) {
        for (var d = 0; d < analysisData.individualResults.length; d++) {
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
            
            // Show compound path warnings prominently
            if (docResult.compoundPathWarnings && docResult.compoundPathWarnings.length > 0) {
                docReport.push("");
                docReport.push("*** CRITICAL: COMPOUND PATH ERRORS ***");
                for (var w = 0; w < docResult.compoundPathWarnings.length; w++) {
                    docReport.push(docResult.compoundPathWarnings[w].text);
                }
            }
            
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
                    var paddedCount = (countStr + "      ").substr(0, 6);
                    docReport.push(paddedCount + "\t" + size);
                }
            } else {
                docReport.push("No CutThrough paths found");
            }
            
            // Show ALL image details in individual document tabs
            if (docResult.allImageDetails && docResult.allImageDetails.length > 0) {
                docReport.push("");
                docReport.push("ALL IMAGE DETAILS");
                for (var img = 0; img < docResult.allImageDetails.length; img++) {
                    docReport.push(docResult.allImageDetails[img].text);
                }
            }
            
            // Add low-res image details if available
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
                    totalCutThroughPaths, totalCompoundPaths, allCompoundPathWarnings,
                    includePPI, individualDocumentResults) {
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
    
    // Add compound path warnings prominently at top
    if (totalCompoundPaths > 0) {
        report.push("*** CRITICAL PRODUCTION ERROR ***");
        report.push("COMPOUND PATHS WITH CUTTHROUGH COLOR DETECTED!");
        report.push("Found " + totalCompoundPaths + " compound path(s) with CutThrough color");
        report.push("This WILL cause cutting machine failures!");
        report.push("MUST BE FIXED before production!");
        report.push("");
        report.push("Compound Path Details:");
        for (var i = 0; i < allCompoundPathWarnings.length; i++) {
            report.push("  " + allCompoundPathWarnings[i].document + ": " + allCompoundPathWarnings[i].text);
        }
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
    if (totalCompoundPaths > 0) {
        report.push("Compound Paths with CutThrough Color: " + totalCompoundPaths + " ** CRITICAL ERROR **");
    }
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
        report.push("");
        report.push("Document: " + docResult.name);
        if (docResult.scaleFactor && docResult.scaleFactor !== 1) {
            report.push("  ** Large Canvas (Scale Factor: " + docResult.scaleFactor + ") **");
        }
        
        // Show compound path warnings for this document
        if (docResult.compoundPathWarnings && docResult.compoundPathWarnings.length > 0) {
            report.push("  ** COMPOUND PATH ERRORS DETECTED **");
            for (var w = 0; w < docResult.compoundPathWarnings.length; w++) {
                report.push("    " + docResult.compoundPathWarnings[w].text);
            }
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
    
    // Add all image details to report
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