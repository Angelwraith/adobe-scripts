#target illustrator

/*@METADATA{
  "name": "Rearrange Registration Dots",
  "description": "Rearranges 5 selected registration dot groups to fit empty space on active artboard",
  "version": "1.0",
  "target": "illustrator",
  "tags": ["registration", "dots", "artboard"]
}@END_METADATA*/

// Main execution
try {
    if (app.documents.length === 0) {
        alert("Please open a document first.");
    } else {
        var doc = app.activeDocument;
        
        // Check if selection exists and has exactly 5 groups
        if (doc.selection.length !== 5) {
            alert("Please select exactly 5 registration dot groups.\n\nYou have selected: " + doc.selection.length + " items");
        } else {
            // Verify all selected items are groups
            var allGroups = true;
            for (var i = 0; i < doc.selection.length; i++) {
                if (doc.selection[i].typename !== "GroupItem") {
                    allGroups = false;
                    break;
                }
            }
            
            if (!allGroups) {
                alert("All selected items must be groups.\n\nPlease select the 5 registration dot groups.");
            } else {
                var success = rearrangeRegistrationDots(doc);
                if (success) {
                    alert("Registration dots rearranged successfully!");
                }
            }
        }
    }
} catch (e) {
    alert("Error: " + e.toString() + "\n\nLine: " + e.line);
}

function showPlacementDialog() {
    var dialog = new Window("dialog", "Registration Dot Placement");
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.spacing = 15;
    dialog.margins = 30;
    
    dialog.add("statictext", undefined, "Where should the registration dots be placed?");
    
    var radioGroup = dialog.add("group");
    radioGroup.orientation = "column";
    radioGroup.alignChildren = "left";
    radioGroup.spacing = 10;
    
    var topBottomRadio = radioGroup.add("radiobutton", undefined, "Top/Bottom (for landscape artboards)");
    var leftRightRadio = radioGroup.add("radiobutton", undefined, "Left/Right (for portrait artboards)");
    
    topBottomRadio.value = true;
    
    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.spacing = 10;
    buttonGroup.alignment = "center";
    
    var okBtn = buttonGroup.add("button", undefined, "OK");
    var cancelBtn = buttonGroup.add("button", undefined, "Cancel");
    
    var result = null;
    
    okBtn.onClick = function() {
        if (topBottomRadio.value) {
            result = "TopBottom";
        } else {
            result = "LeftRight";
        }
        dialog.close();
    };
    
    cancelBtn.onClick = function() {
        result = null;
        dialog.close();
    };
    
    dialog.show();
    return result;
}

function rearrangeRegistrationDots(doc) {
    // Get the active artboard
    var artboardIndex = doc.artboards.getActiveArtboardIndex();
    var artboard = doc.artboards[artboardIndex];
    var bounds = artboard.artboardRect;
    
    var left = bounds[0];
    var top = bounds[1];
    var right = bounds[2];
    var bottom = bounds[3];
    
    // Store the selected groups
    var dotGroups = [];
    for (var i = 0; i < doc.selection.length; i++) {
        dotGroups.push(doc.selection[i]);
    }
    
    // Calculate the average radius from existing dots
    var dotRadius = calculateDotRadius(dotGroups);
    
    // Prompt user for placement choice
    var placement = showPlacementDialog();
    if (!placement) {
        // User cancelled
        return false;
    }
    
    // Calculate new positions based on placement
    var dotPositions = [];
    
    if (placement === "TopBottom") {
        var effectiveLeft = left + dotRadius;
        var effectiveRight = right - dotRadius;
        var effectiveWidth = effectiveRight - effectiveLeft;
        var spacing = effectiveWidth / 4;
        
        // Create 5 positions
        for (var i = 0; i < 5; i++) {
            var x = effectiveLeft + (i * spacing);
            var y = top - dotRadius;
            dotPositions.push({x: x, y: y});
        }
        
        // Move dots 1 and 3 to bottom
        dotPositions[1].y = bottom + dotRadius;
        dotPositions[3].y = bottom + dotRadius;
        
    } else {
        // LeftRight placement
        var effectiveTop = top - dotRadius;
        var effectiveBottom = bottom + dotRadius;
        var effectiveHeight = effectiveTop - effectiveBottom;
        var spacing = effectiveHeight / 4;
        
        // Create 5 positions
        for (var i = 0; i < 5; i++) {
            var x = left + dotRadius;
            var y = effectiveTop - (i * spacing);
            dotPositions.push({x: x, y: y});
        }
        
        // Move dots 1 and 3 to right
        dotPositions[1].x = right - dotRadius;
        dotPositions[3].x = right - dotRadius;
    }
    
    // Move each dot group to its new position
    for (var i = 0; i < dotGroups.length; i++) {
        moveDotGroupToPosition(dotGroups[i], dotPositions[i].x, dotPositions[i].y);
    }
    
    return true;
}

function calculateDotRadius(dotGroups) {
    // Calculate radius from the first dot group
    // Assuming the outer circle is the bounding box of the group
    var firstGroup = dotGroups[0];
    var groupBounds = firstGroup.geometricBounds;
    
    // geometricBounds: [left, top, right, bottom]
    var groupWidth = groupBounds[2] - groupBounds[0];
    var groupHeight = groupBounds[1] - groupBounds[3];
    
    // Use the average of width and height to get diameter, then divide by 2
    var diameter = (groupWidth + groupHeight) / 2;
    var radius = diameter / 2;
    
    return radius;
}

function moveDotGroupToPosition(dotGroup, centerX, centerY) {
    // Get the current center of the group
    var groupBounds = dotGroup.geometricBounds;
    var currentCenterX = (groupBounds[0] + groupBounds[2]) / 2;
    var currentCenterY = (groupBounds[1] + groupBounds[3]) / 2;
    
    // Calculate the offset needed
    var deltaX = centerX - currentCenterX;
    var deltaY = centerY - currentCenterY;
    
    // Move the group by translating it
    dotGroup.translate(deltaX, deltaY);
}
