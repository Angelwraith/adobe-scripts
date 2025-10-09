#target illustrator

/*
@METADATA
{
  "name": "File Name Sorter",
  "description": "Organize Selected Objects by Sequential File Names",
  "version": "1.0",
  "target": "illustrator",
  "tags": ["Sorter", "processors", "name"]
}
@END_METADATA
*/

function getObjectDisplayName(obj) {
    try {
        // Check for PlacedItem (linked files)
        if (obj.typename === "PlacedItem") {
            if (obj.file && obj.file.name) {
                var fileName = decodeURI(obj.file.name);
                // Remove file extension using lastIndexOf method
                var lastDot = fileName.lastIndexOf('.');
                if (lastDot > 0) {
                    return fileName.substring(0, lastDot);
                }
                return fileName;
            }
        }
        
        // Check for RasterItem
        if (obj.typename === "RasterItem") {
            if (obj.file && obj.file.name) {
                var fileName = decodeURI(obj.file.name);
                var lastDot = fileName.lastIndexOf('.');
                if (lastDot > 0) {
                    return fileName.substring(0, lastDot);
                }
                return fileName;
            }
        }
        
        // Check for GroupItem - might contain linked items
        if (obj.typename === "GroupItem" && obj.pageItems.length > 0) {
            // Look for the first PlacedItem in the group
            for (var i = 0; i < obj.pageItems.length; i++) {
                if (obj.pageItems[i].typename === "PlacedItem" && obj.pageItems[i].file) {
                    var fileName = decodeURI(obj.pageItems[i].file.name);
                    var lastDot = fileName.lastIndexOf('.');
                    if (lastDot > 0) {
                        return fileName.substring(0, lastDot);
                    }
                    return fileName;
                }
            }
        }
        
        // Fall back to object name
        if (obj.name && obj.name.length > 0) {
            return obj.name;
        }
        
    } catch (e) {
        // If any error occurs, return null
    }
    
    return null;
}

function listSelectedObjectNames() {
    if (app.documents.length === 0) {
        alert("Please open a document first.");
        return;
    }
    
    var doc = app.activeDocument;
    var selection = doc.selection;
    
    if (selection.length === 0) {
        alert("Please select some objects first.");
        return;
    }
    
    var namesList = "Selected object details:\n\n";
    for (var i = 0; i < selection.length; i++) {
        var obj = selection[i];
        var displayName = getObjectDisplayName(obj);
        var type = obj.typename || "unknown";
        
        namesList += (i + 1) + ". Type: " + type + "\n";
        namesList += "   Display Name: " + (displayName || "[no name found]") + "\n";
        namesList += "   Object Name: " + (obj.name || "[none]") + "\n";
        
        // Add parsing results for debugging
        if (displayName) {
            var baseName = extractBaseName(displayName);
            var seqNumber = extractSequenceNumber(displayName);
            namesList += "   Parsed Base: " + (baseName || "[none]") + "\n";
            namesList += "   Parsed Number: " + (seqNumber || "[none]") + "\n";
        }
        
        // Try to get file info for debugging
        try {
            if (obj.typename === "PlacedItem" && obj.file) {
                namesList += "   File: " + decodeURI(obj.file.name) + "\n";
            }
        } catch (e) {
            namesList += "   File: [error accessing file]\n";
        }
        
        namesList += "\n";
    }
    
    alert(namesList);
}

function organizeBySequentialNames() {
    // Check if there's an active document
    if (app.documents.length === 0) {
        alert("Please open a document first.");
        return;
    }
    
    var doc = app.activeDocument;
    var selection = doc.selection;
    
    // Check if anything is selected
    if (selection.length === 0) {
        alert("Please select some objects first.");
        return;
    }
    
    // Extract objects with valid names (from linked files or object names)
    var namedObjects = [];
    for (var i = 0; i < selection.length; i++) {
        var obj = selection[i];
        var objectName = getObjectDisplayName(obj);
        
        if (objectName && objectName.length > 0) {
            namedObjects.push({
                object: obj,
                name: objectName,
                bounds: obj.geometricBounds
            });
        }
    }
    
    if (namedObjects.length === 0) {
        alert("No objects with usable names found in selection.\n\nTry running the debug function first to see what names are being detected.");
        return;
    }
    
    // Group objects by their base name pattern
    var groups = groupByBaseName(namedObjects);
    
    // Check if we found any valid groups
    var groupCount = 0;
    for (var groupName in groups) {
        groupCount++;
    }
    
    if (groupCount === 0) {
        alert("No valid sequential naming patterns found.\n\nMake sure your files end with numbers (e.g., 'filename1', 'filename2', etc.)");
        return;
    }
    
    // Sort and arrange each group
    arrangeGroups(groups);
    
    // Now use Illustrator's alignment tools for vertical centering only
    var doc = app.activeDocument;
    doc.selection = [];
    
    // Re-select all the organized objects in order
    var allObjects = [];
    for (var groupName in groups) {
        var group = groups[groupName];
        for (var i = 0; i < group.length; i++) {
            allObjects.push(group[i]);
        }
    }
    
    // Sort objects by sequence number
    allObjects.sort(function(a, b) {
        return a.sequenceNumber - b.sequenceNumber;
    });
    
    // Select all objects
    for (var i = 0; i < allObjects.length; i++) {
        allObjects[i].object.selected = true;
    }
    
    // Use vertical align center only
    if (allObjects.length > 0) {
        try {
            app.executeMenuCommand("Vertical Align Center");
        } catch (e) {
            // Silent error handling - no popup
        }
    }
}

function groupByBaseName(objects) {
    var groups = {};
    
    for (var i = 0; i < objects.length; i++) {
        var obj = objects[i];
        var baseName = extractBaseName(obj.name);
        var sequenceNumber = extractSequenceNumber(obj.name);
        
        if (baseName !== null && sequenceNumber !== null) {
            if (!groups[baseName]) {
                groups[baseName] = [];
            }
            
            groups[baseName].push({
                object: obj.object,
                name: obj.name,
                baseName: baseName,
                sequenceNumber: sequenceNumber,
                bounds: obj.bounds
            });
        }
    }
    
    // Sort each group by sequence number
    for (var groupName in groups) {
        groups[groupName].sort(function(a, b) {
            return a.sequenceNumber - b.sequenceNumber;
        });
    }
    
    return groups;
}

function extractBaseName(name) {
    // Look specifically for pattern ending with P + numbers (like P1, P2, P7, etc.)
    // This handles your specific naming pattern better
    var i = name.length - 1;
    var hasFoundDigit = false;
    
    // Go backwards until we find the last non-digit character
    while (i >= 0 && name.charAt(i) >= '0' && name.charAt(i) <= '9') {
        hasFoundDigit = true;
        i--;
    }
    
    if (hasFoundDigit && i >= 0) {
        return name.substring(0, i + 1);
    }
    
    return null;
}

function extractSequenceNumber(name) {
    // Extract trailing numbers from the end of the name
    var i = name.length - 1;
    var numberStr = "";
    
    while (i >= 0 && name.charAt(i) >= '0' && name.charAt(i) <= '9') {
        numberStr = name.charAt(i) + numberStr;
        i--;
    }
    
    if (numberStr.length > 0) {
        return parseInt(numberStr, 10);
    }
    
    return null;
}

function arrangeGroups(groups) {
    // Collect all objects from all groups
    var allObjects = [];
    for (var groupName in groups) {
        var group = groups[groupName];
        for (var i = 0; i < group.length; i++) {
            allObjects.push(group[i]);
        }
    }
    
    // Sort all objects by sequence number
    allObjects.sort(function(a, b) {
        return a.sequenceNumber - b.sequenceNumber;
    });
    
    if (allObjects.length === 0) return;
    
    // Use the position of the first object as the starting point
    var firstObj = allObjects[0].object;
    var firstBounds = firstObj.geometricBounds;
    var startX = firstBounds[1]; // Left edge of first object
    var fixedY = firstBounds[0]; // Top edge of first object - ALL objects get this Y
    
    var gapSpacing = 10 * 72; // 10 inches converted to points
    var currentX = startX;
    
    // Position each object
    for (var i = 0; i < allObjects.length; i++) {
        var obj = allObjects[i].object;
        
        // Calculate target position
        var targetX = currentX;
        var targetY = fixedY; // Same Y for ALL objects
        
        // Get current position
        var bounds = obj.geometricBounds; // [top, left, bottom, right]
        var currentLeft = bounds[1];
        var currentTop = bounds[0];
        var currentRight = bounds[3];
        
        // Calculate actual width of this object
        var objectWidth = currentRight - bounds[1];
        
        // Calculate movement needed
        var deltaX = targetX - currentLeft;
        var deltaY = targetY - currentTop;
        
        // Move the object
        obj.translate(deltaX, deltaY);
        
        // Update currentX for next object: current position + object width + gap
        currentX = targetX + objectWidth + gapSpacing;
    }
}

// Main execution
try {
    // For normal operation:
    organizeBySequentialNames();
    
    // For debugging only, comment out line above and uncomment this:
    // listSelectedObjectNames();
} catch (error) {
    alert("Error: " + error.message);
}