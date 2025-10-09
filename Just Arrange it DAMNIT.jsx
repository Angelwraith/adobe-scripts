/*
@METADATA
{
  "name": "Just Arrange it DAMNIT",
  "description": "Poorly arranges objects on artboards.",
  "version": "2.0",
  "target": "illustrator",
  "tags": ["Arrange", "pack", "processors"]
}
@END_METADATA
*/

// Adobe Illustrator Script: Arrange Selected Groups on Artboard
// This script rearranges selected groups onto the active artboard with user-defined spacing
// Prioritizes fitting maximum groups by rotating as needed, selects groups that don't fit

(function() {
    var startTime = new Date().getTime(); // Start timer
    
    // Check if there's an active document
    if (app.documents.length === 0) {
        alert("Please open a document first.");
        return;
    }
    
    var doc = app.activeDocument;
    var selection = doc.selection;
    
    // Filter selection to only include groups
    var groups = [];
    for (var i = 0; i < selection.length; i++) {
        if (selection[i].typename === "GroupItem") {
            groups.push(selection[i]);
        }
    }
    
    // Check if any groups are selected
    if (groups.length === 0) {
        alert("Please select one or more groups to arrange.");
        return;
    }
    
    // Create modal dialog for user input
    var dialog = new Window("dialog", "Arrange Groups Settings");
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.spacing = 10;
    dialog.margins = 16;
    
    // Orientation selection
    var orientationPanel = dialog.add("panel", undefined, "Reg dot sides:");
    orientationPanel.orientation = "row";
    orientationPanel.alignChildren = "left";
    orientationPanel.spacing = 10;
    orientationPanel.margins = 10;
    
    var topBottomBtn = orientationPanel.add("radiobutton", undefined, "Top/Bottom");
    var leftRightBtn = orientationPanel.add("radiobutton", undefined, "Left/Right");
    topBottomBtn.value = true; // Default selection
    
    // Buttons
    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignment = "center";
    buttonGroup.spacing = 10;
    
    var okButton = buttonGroup.add("button", undefined, "OK");
    var cancelButton = buttonGroup.add("button", undefined, "Cancel");
    
    okButton.onClick = function() {
        dialog.close(1);
    };
    
    cancelButton.onClick = function() {
        dialog.close(0);
    };
    
    // Show dialog and get results
    var result = dialog.show();
    if (result === 0) {
        return; // User cancelled
    }
    
    var useTopBottom = topBottomBtn.value;
    
    // No spacing - groups will be placed edge-to-edge with slight overlap tolerance
    var documentScaleFactor = app.activeDocument.scaleFactor;
    var oneInchInPoints = 72 / documentScaleFactor;
    var overlapTolerance = (0.1 * 72) / documentScaleFactor; // Allow 0.1" overlap
    
    // Get the active artboard
    var artboard = doc.artboards[doc.artboards.getActiveArtboardIndex()];
    var artboardRect = artboard.artboardRect;
    
    // Artboard dimensions (left, top, right, bottom)
    var artboardLeft = artboardRect[0];
    var artboardTop = artboardRect[1];
    var artboardRight = artboardRect[2];
    var artboardBottom = artboardRect[3];
    var artboardWidth = artboardRight - artboardLeft;
    var artboardHeight = artboardTop - artboardBottom;
    
    // Calculate available area based on hotdog/hamburger selection
    var availableLeft, availableTop, availableRight, availableBottom;
    var availableWidth, availableHeight;
    
    if (useTopBottom) {
        // Top/Bottom: exclude 1 inch from top and bottom
        availableLeft = artboardLeft;
        availableTop = artboardTop - oneInchInPoints;
        availableRight = artboardRight;
        availableBottom = artboardBottom + oneInchInPoints;
        availableWidth = artboardWidth;
        availableHeight = artboardHeight - (2 * oneInchInPoints);
    } else {
        // Left/Right: exclude 1 inch from left and right
        availableLeft = artboardLeft + oneInchInPoints;
        availableTop = artboardTop;
        availableRight = artboardRight - oneInchInPoints;
        availableBottom = artboardBottom;
        availableWidth = artboardWidth - (2 * oneInchInPoints);
        availableHeight = artboardHeight;
    }
    
    // Prepare group data with both orientations
    var groupData = [];
    
    for (var i = 0; i < groups.length; i++) {
        var group = groups[i];
        var bounds = group.visibleBounds;
        var originalWidth = bounds[2] - bounds[0];
        var originalHeight = bounds[1] - bounds[3];
        
        groupData.push({
            group: group,
            originalBounds: bounds,
            normalOrientation: {
                width: originalWidth,
                height: originalHeight,
                rotated: false
            },
            rotatedOrientation: {
                width: originalHeight,
                height: originalWidth,
                rotated: true
            },
            area: originalWidth * originalHeight,
            bestOrientation: null,
            placed: false,
            originalLeft: bounds[0],
            originalTop: bounds[1]
        });
    }
    
    // Sort groups by position: top to bottom, then left to right
    groupData.sort(function(a, b) {
        var tolerance = 10; // 10 points tolerance for "same row"
        
        // If groups are roughly on the same horizontal level (within tolerance)
        if (Math.abs(a.originalTop - b.originalTop) <= tolerance) {
            return a.originalLeft - b.originalLeft; // Sort by left position (left to right)
        } else {
            return b.originalTop - a.originalTop; // Sort by top position (top to bottom)
        }
    });
    
    // Try to place groups using a greedy bin packing approach
    var placedGroups = [];
    var unplacedGroups = [];
    var occupiedRects = []; // Track occupied spaces
    var consecutiveFailures = 0;
    var maxConsecutiveFailures = Math.min(10, Math.ceil(groups.length * 0.1)); // Stop after 10 failures or 10% of groups fail consecutively
    
    // Detect existing groups on the active layer that are within the artboard
    function detectExistingGroups() {
        var existingGroups = [];
        
        function searchItems(items) {
            for (var i = 0; i < items.length; i++) {
                var item = items[i];
                if (item.typename === "GroupItem") {
                    // Skip if this group is one of the selected groups we're trying to place
                    var isSelectedGroup = false;
                    for (var j = 0; j < groups.length; j++) {
                        if (item === groups[j]) {
                            isSelectedGroup = true;
                            break;
                        }
                    }
                    
                    if (!isSelectedGroup) {
                        try {
                            var bounds = item.visibleBounds;
                            
                            // Check if group overlaps with the artboard area
                            if (!(bounds[2] < artboardLeft || bounds[0] > artboardRight || 
                                  bounds[3] > artboardTop || bounds[1] < artboardBottom)) {
                                existingGroups.push({
                                    left: bounds[0],
                                    top: bounds[1],
                                    right: bounds[2],
                                    bottom: bounds[3]
                                });
                            }
                        } catch (e) {
                            // Skip items that can't be measured
                        }
                    }
                }
            }
        }
        
        function searchLayers(layers) {
            for (var i = 0; i < layers.length; i++) {
                searchItems(layers[i].pageItems);
                if (layers[i].layers.length > 0) {
                    searchLayers(layers[i].layers);
                }
            }
        }
        
        // Search the active layer
        searchItems(doc.activeLayer.pageItems);
        searchLayers(doc.activeLayer.layers);
        
        return existingGroups;
    }
    
    // Add existing groups to occupied rectangles
    var existingGroups = detectExistingGroups();
    occupiedRects = occupiedRects.concat(existingGroups);
    
    function rectanglesOverlap(rect1, rect2) {
        // Allow up to 0.1" overlap for better edge-to-edge placement
        return !(rect1.right - overlapTolerance <= rect2.left ||
                 rect2.right - overlapTolerance <= rect1.left ||
                 rect1.bottom + overlapTolerance >= rect2.top ||
                 rect2.bottom + overlapTolerance >= rect1.top);
    }
    
    function canPlaceAt(x, y, width, height) {
        // Check available area boundaries (respects reg dot setting)
        if (x < availableLeft ||
            y - height < availableBottom ||
            x + width > availableRight ||
            y > availableTop) {
            return false;
        }
        
        var newRect = {
            left: x,
            top: y,
            right: x + width,
            bottom: y - height
        };
        
        // Check overlap with existing groups
        for (var i = 0; i < occupiedRects.length; i++) {
            if (rectanglesOverlap(newRect, occupiedRects[i])) {
                return false;
            }
        }
        
        return true;
    }
    
    function findBestPosition(width, height) {
        var stepSize = Math.max(5, 5 / documentScaleFactor);
        
        // Early exit if the group is obviously too big for the available space
        if (width > availableWidth || height > availableHeight) {
            return null;
        }
        
        for (var y = availableTop; y >= availableBottom + height; y -= stepSize) {
            for (var x = availableLeft; x <= availableRight - width; x += stepSize) {
                if (canPlaceAt(x, y, width, height)) {
                    return { x: x, y: y };
                }
            }
        }
        return null;
    }
    
    // Process each group
    for (var i = 0; i < groupData.length; i++) {
        var data = groupData[i];
        var bestPosition = null;
        var bestOrientation = null;
        
        // Try normal orientation first
        var normalPos = findBestPosition(data.normalOrientation.width, data.normalOrientation.height);
        if (normalPos) {
            bestPosition = normalPos;
            bestOrientation = data.normalOrientation;
        }
        
        // Try rotated orientation if normal didn't work or if rotated fits better
        var rotatedPos = findBestPosition(data.rotatedOrientation.width, data.rotatedOrientation.height);
        if (rotatedPos && (!bestPosition ||
            rotatedPos.y > bestPosition.y || // Higher position is better
            (rotatedPos.y === bestPosition.y && rotatedPos.x < bestPosition.x))) { // Leftmost if same height
            bestPosition = rotatedPos;
            bestOrientation = data.rotatedOrientation;
        }
        
        if (bestPosition && bestOrientation) {
            // Place the group
            data.bestOrientation = bestOrientation;
            data.position = bestPosition;
            data.placed = true;
            consecutiveFailures = 0; // Reset failure counter
            
            // Apply rotation if needed
            if (bestOrientation.rotated) {
                var rotationMatrix = app.getRotationMatrix(-90);
                data.group.transform(rotationMatrix);
            }
            
            // Move to position
            var currentBounds = data.group.visibleBounds;
            var deltaX = bestPosition.x - currentBounds[0];
            var deltaY = bestPosition.y - currentBounds[1];
            
            var translateMatrix = app.getTranslationMatrix(deltaX, deltaY);
            data.group.transform(translateMatrix);
            
            // Get final bounds and add to occupied rectangles
            var finalBounds = data.group.visibleBounds;
            occupiedRects.push({
                left: finalBounds[0],
                top: finalBounds[1],
                right: finalBounds[2],
                bottom: finalBounds[3]
            });
            
            placedGroups.push(data);
        } else {
            // Couldn't place this group
            consecutiveFailures++;
            unplacedGroups.push(data);
            
            // More lenient early termination: only stop if we've failed on many consecutive groups
            // AND we've already placed at least some groups successfully
            if (consecutiveFailures >= Math.max(maxConsecutiveFailures, 20) && placedGroups.length > 0) {
                // Check if any of the remaining groups are significantly smaller and might still fit
                var hasSmallRemainingGroups = false;
                var averagePlacedArea = 0;
                
                // Calculate average area of successfully placed groups
                for (var p = 0; p < placedGroups.length; p++) {
                    averagePlacedArea += placedGroups[p].area;
                }
                averagePlacedArea = averagePlacedArea / placedGroups.length;
                
                // Check if any remaining groups are significantly smaller
                for (var r = i + 1; r < groupData.length; r++) {
                    if (groupData[r].area < averagePlacedArea * 0.5) {
                        hasSmallRemainingGroups = true;
                        break;
                    }
                }
                
                // Only terminate early if there are no small remaining groups
                if (!hasSmallRemainingGroups) {
                    // Add all remaining groups to unplaced
                    for (var j = i + 1; j < groupData.length; j++) {
                        unplacedGroups.push(groupData[j]);
                    }
                    break;
                }
            }
        }
    }
    
    // Select groups that couldn't fit
    doc.selection = null; // Clear selection first
    if (unplacedGroups.length > 0) {
        var newSelection = [];
        for (var i = 0; i < unplacedGroups.length; i++) {
            newSelection.push(unplacedGroups[i].group);
        }
        doc.selection = newSelection;
    }
    
    // Single refresh at the end (performance improvement)
    app.redraw();
    
    // Count rotated groups
    var rotatedCount = 0;
    for (var i = 0; i < placedGroups.length; i++) {
        if (placedGroups[i].bestOrientation && placedGroups[i].bestOrientation.rotated) {
            rotatedCount++;
        }
    }
    
    // Calculate execution time
    var endTime = new Date().getTime();
    var executionTime = ((endTime - startTime) / 1000).toFixed(2);
    var timePerGroup = (executionTime / groups.length).toFixed(3);
    
    // Create status message
    var orientationText = useTopBottom ? "Top/Bottom" : "Left/Right";
    var message = "Arranged " + placedGroups.length + " of " + groups.length + " groups.";
    message += "\nUsing " + orientationText + " Reg dots orientation.";
    message += "\nExecution time: " + executionTime + " seconds (" + timePerGroup + " sec/group).";
    
    if (existingGroups.length > 0) {
        message += "\nAvoided " + existingGroups.length + " existing group(s) on the artboard.";
    }
    
    if (rotatedCount > 0) {
        message += "\nRotated " + rotatedCount + " group(s) 90 degrees for better fit.";
    }
    
    if (unplacedGroups.length > 0) {
        message += "\n" + unplacedGroups.length + " group(s) could not fit and are now selected.";
    }
    
    alert(message);
})();