// Illustrator Script: Sort Cuts Inside to Outside
// Assigns spot colors based on path nesting level

// Create or get spot colors
function getOrCreateSpotColor(name, c, m, y, k) {
    // Check if spot already exists
    var spots = app.activeDocument.spots;
    for (var i = 0; i < spots.length; i++) {
        if (spots[i].name === name) {
            return spots[i];
        }
    }
    
    // Create new spot color
    var spot = app.activeDocument.spots.add();
    spot.name = name;
    spot.colorType = ColorModel.SPOT;
    var color = new CMYKColor();
    color.cyan = c;
    color.magenta = m;
    color.yellow = y;
    color.black = k;
    spot.color = color;
    
    return spot;
}

// Create spot colors from your swatches
var insideColor = getOrCreateSpotColor("CutThrough1-Inside", 100, 0, 0, 0);
var outsideColor = getOrCreateSpotColor("CutThrough2-Outside", 0, 100, 0, 0);

// Check if a point is inside a path
function isPointInPath(point, path) {
    if (!path.closed) return false;
    
    var bounds = path.geometricBounds;
    if (point[0] < bounds[0] || point[0] > bounds[2] ||
        point[1] > bounds[1] || point[1] < bounds[3]) {
        return false;
    }
    
    // Use Illustrator's built-in hit test
    try {
        var testPath = app.activeDocument.pathItems.add();
        testPath.setEntirePath([[point[0], point[1]]]);
        
        var result = false;
        if (path.closed) {
            // Check if point is inside using area method
            var x = point[0], y = point[1];
            var pathPoints = path.pathPoints;
            var inside = false;
            
            for (var i = 0, j = pathPoints.length - 1; i < pathPoints.length; j = i++) {
                var xi = pathPoints[i].anchor[0], yi = pathPoints[i].anchor[1];
                var xj = pathPoints[j].anchor[0], yj = pathPoints[j].anchor[1];
                
                var intersect = ((yi > y) != (yj > y)) && 
                    (x < (xj - xi) * (y - yi) / (yj - yi) + xi);
                if (intersect) inside = !inside;
            }
            result = inside;
        }
        
        testPath.remove();
        return result;
    } catch (e) {
        return false;
    }
}

// Get all closed paths from selection or document
function getAllClosedPaths() {
    var paths = [];
    var items = app.activeDocument.selection.length > 0 ? 
                app.activeDocument.selection : 
                app.activeDocument.pathItems;
    
    for (var i = 0; i < items.length; i++) {
        if (items[i].typename === "PathItem" && items[i].closed) {
            paths.push(items[i]);
        }
    }
    return paths;
}

// Determine nesting level of each path
function getNestingLevel(path, allPaths) {
    var level = 0;
    var testPoint = path.pathPoints[0].anchor;
    
    for (var i = 0; i < allPaths.length; i++) {
        if (allPaths[i] === path) continue;
        if (isPointInPath(testPoint, allPaths[i])) {
            level++;
        }
    }
    return level;
}

// Apply spot color to path
function applySpotColor(path, spotColor) {
    var spotColorRef = new SpotColor();
    spotColorRef.spot = spotColor;
    
    if (path.filled) {
        path.fillColor = spotColorRef;
    }
    if (path.stroked) {
        path.strokeColor = spotColorRef;
    }
}

// Main function
function main() {
    if (app.documents.length === 0) {
        alert("Please open a document first.");
        return;
    }
    
    var paths = getAllClosedPaths();
    
    if (paths.length === 0) {
        alert("No closed paths found. Please select paths or ensure document contains closed paths.");
        return;
    }
    
    var pathsWithLevels = [];
    
    // Calculate nesting level for each path
    for (var i = 0; i < paths.length; i++) {
        var level = getNestingLevel(paths[i], paths);
        pathsWithLevels.push({
            path: paths[i],
            level: level
        });
    }
    
    // Sort by level (inside to outside)
    pathsWithLevels.sort(function(a, b) {
        return b.level - a.level;
    });
    
    // Apply colors: odd levels = inside, even levels = outside
    var insideCount = 0;
    var outsideCount = 0;
    
    for (var i = 0; i < pathsWithLevels.length; i++) {
        var item = pathsWithLevels[i];
        if (item.level % 2 === 1) {
            // Odd level = inside cut
            applySpotColor(item.path, insideColor);
            insideCount++;
        } else {
            // Even level = outside cut
            applySpotColor(item.path, outsideColor);
            outsideCount++;
        }
    }
    
    alert("Processing complete!\n" +
          "Inside cuts: " + insideCount + " (CutThrough1-Inside)\n" +
          "Outside cuts: " + outsideCount + " (CutThrough2-Outside)");
}

// Run the script
main();