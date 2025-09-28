/*
@METADATA
{
  "name": "Group Un-Rotator",
  "description": "Un-Rotate a Group by a Baseline Reference",
  "version": "1.0",
  "target": "illustrator",
  "tags": ["Un-Rotate", "group", "transform"]
}
@END_METADATA
*/

#target illustrator

function main() {
    try {
        if (app.documents.length == 0) {
            return; // Silent fail
        }
        
        var doc = app.activeDocument;
        var selection = doc.selection;
        
        if (selection.length == 0) {
            alert("Please select your artwork first.");
            return;
        }
        
        // Check if selection contains a simple line (baseline)
        var baselinePath = null;
        var artworkItems = [];
        
        for (var i = 0; i < selection.length; i++) {
            var item = selection[i];
            
            // Check if this is a simple 2-point line
            if (item.typename == "PathItem" && 
                item.pathPoints.length == 2 && 
                !item.closed) {
                baselinePath = item;
            } else {
                artworkItems.push(item);
            }
        }
        
        if (baselinePath && artworkItems.length > 0) {
            // We have both baseline and artwork - perform rotation
            performBaselineRotation(baselinePath, artworkItems);
        } else if (artworkItems.length > 0 && !baselinePath) {
            // Only artwork selected - store for next run
            alert("\n1. With the line tool draw a baseline on the target group\n2. Select BOTH the group and line\n3. Run this script again");
        } else {
            alert("Please select your artwork and/or draw a baseline reference line.");
        }
        
    } catch (e) {
        alert("Error: " + e.message);
    }
}

function performBaselineRotation(baselinePath, artworkItems) {
    try {
        // Calculate the baseline angle
        var startPoint = baselinePath.pathPoints[0].anchor;
        var endPoint = baselinePath.pathPoints[1].anchor;
        
        var deltaX = endPoint[0] - startPoint[0];
        var deltaY = endPoint[1] - startPoint[1];
        
        // Calculate angle in degrees
        var angleRad = Math.atan2(deltaY, deltaX);
        var angleDeg = angleRad * (180 / Math.PI);
        
        // Delete the baseline
        baselinePath.remove();
        
        // Create a temporary group to rotate everything together
        var tempGroup = app.activeDocument.groupItems.add();
        
        // Move all artwork items into the group
        for (var i = 0; i < artworkItems.length; i++) {
            artworkItems[i].move(tempGroup, ElementPlacement.INSIDE);
        }
        
        // Rotate the entire group by negative of baseline angle
        var rotationAngle = -angleDeg;
        tempGroup.rotate(rotationAngle);
        
        // Ungroup the items back to their original state
        var ungroupedItems = [];
        for (var j = tempGroup.pageItems.length - 1; j >= 0; j--) {
            var item = tempGroup.pageItems[j];
            item.move(app.activeDocument, ElementPlacement.PLACEATEND);
            ungroupedItems.push(item);
        }
        
        // Remove the temporary group
        tempGroup.remove();
        
        // Reselect the artwork
        app.activeDocument.selection = ungroupedItems;
        
        // Force redraw
        app.redraw();
                
    } catch (e) {
        alert("Rotation error: " + e.message);
    }
}

// Run the script
main();