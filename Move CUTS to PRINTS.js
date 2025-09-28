/*
@METADATA
{
  "name": "CUT Mover",
  "description": "Moves CUT Files to PRINT File Positions",
  "version": "1.0",
  "target": "illustrator",
  "tags": ["CUT", "mover", "utility"]
}
@END_METADATA
*/

(function() {
    // Check if we have an active document
    if (app.documents.length === 0) {
        alert("No active document found. Please open a document with linked files.");
        return;
    }

    var doc = app.activeDocument;
    var placedItems = [];
    var printItems = [];
    var cutItems = [];
    var matchedPairs = [];

    // Function to get file name without extension from a path
    function getFileNameWithoutExtension(filePath) {
        var fileName = filePath.substring(filePath.lastIndexOf('/') + 1);
        fileName = fileName.substring(fileName.lastIndexOf('\\') + 1);
        return fileName.substring(0, fileName.lastIndexOf('.'));
    }

    // Function to get the parts before and after PRINT_ or CUT_
    function getFileNameParts(fileName) {
        var printMatch = fileName.match(/^(.*)_PRINT_(.*)$/i);
        if (printMatch) {
            return {
                before: printMatch[1],
                after: printMatch[2],
                type: 'PRINT'
            };
        }
        
        var cutMatch = fileName.match(/^(.*)_CUT_(.*)$/i);
        if (cutMatch) {
            return {
                before: cutMatch[1],
                after: cutMatch[2],
                type: 'CUT'
            };
        }
        
        return null;
    }

    // Function to create a normalized name for matching (before + after)
    function getNormalizedName(fileName) {
        var parts = getFileNameParts(fileName);
        if (parts) {
            return parts.before + "_" + parts.after;
        }
        return fileName;
    }

    // Function to check if name contains PRINT_ or CUT_
    function containsPrint(fileName) {
        return /_PRINT_/i.test(fileName);
    }

    function containsCut(fileName) {
        return /_CUT_/i.test(fileName);
    }

    // Collect all placed items (linked files)
    function collectPlacedItems(container) {
        for (var i = 0; i < container.placedItems.length; i++) {
            placedItems.push(container.placedItems[i]);
        }
        
        // Recursively check groups
        for (var j = 0; j < container.groupItems.length; j++) {
            collectPlacedItems(container.groupItems[j]);
        }
    }

    // Start collecting from the document
    collectPlacedItems(doc);

    if (placedItems.length === 0) {
        alert("No linked files found in the document.");
        return;
    }

    // Categorize items as PRINT or CUT
    for (var i = 0; i < placedItems.length; i++) {
        var item = placedItems[i];
        try {
            var fileName = getFileNameWithoutExtension(item.file.fsName);
            
            if (containsPrint(fileName)) {
                printItems.push({
                    item: item,
                    fileName: fileName,
                    normalizedName: getNormalizedName(fileName)
                });
            } else if (containsCut(fileName)) {
                cutItems.push({
                    item: item,
                    fileName: fileName,
                    normalizedName: getNormalizedName(fileName)
                });
            }
        } catch (e) {
            // Skip items without valid file references
            continue;
        }
    }

    // Find matching pairs
    for (var i = 0; i < printItems.length; i++) {
        var printItem = printItems[i];
        
        for (var j = 0; j < cutItems.length; j++) {
            var cutItem = cutItems[j];
            
            if (printItem.normalizedName === cutItem.normalizedName) {
                matchedPairs.push({
                    print: printItem,
                    cut: cutItem
                });
                break; // Found a match, move to next print item
            }
        }
    }

    if (matchedPairs.length === 0) {
        alert("No matching PRINT/CUT file pairs found.");
        return;
    }

    var processedCount = 0;

    // Process each matched pair
    for (var i = 0; i < matchedPairs.length; i++) {
        var pair = matchedPairs[i];
        var printItem = pair.print.item;
        var cutItem = pair.cut.item;

        try {
            // Move CUT item to same position as PRINT item
            cutItem.position = [printItem.position[0], printItem.position[1]];
            processedCount++;

        } catch (e) {
            // Continue with other pairs if one fails
            continue;
        }
    }

    // Show completion message
    if (processedCount > 0) {
        alert("Script completed successfully!\n\n" +
              "Processed " + processedCount + " PRINT/CUT pairs:\n" +
              "- Moved CUT files to match PRINT file positions");
    } else {
        alert("No pairs could be processed. Please check that:\n" +
              "- Files have matching names with _PRINT_ and _CUT_ patterns\n" +
              "- Everything before and after _PRINT_/_CUT_ is identical");
    }

})();