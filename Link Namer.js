/*
@METADATA
{
  "name": "Link Namer",
  "description": "Name Linked Files for ProdProof Layout",
  "version": "1.1",
  "target": "illustrator",
  "tags": ["ProdProof", "linked", "namer"]
}
@END_METADATA
*/


(function() {
    if (!app.documents.length) {
        alert("Please open a document first.");
        return;
    }
    
    var doc = app.activeDocument;
    var selection = doc.selection;
    
    if (selection.length === 0) {
        alert("Please select one or more linked files.");
        return;
    }
    
    // Get the host file name (without extension)
    var hostFileName = doc.name;
    var dotIndex = hostFileName.lastIndexOf(".");
    if (dotIndex > -1) {
        hostFileName = hostFileName.substring(0, dotIndex);
    }
    
    var processedCount = 0;
    
    for (var i = 0; i < selection.length; i++) {
        var item = selection[i];
        if (item.typename === "PlacedItem") {
            processLinkedFile(item, hostFileName);
            processedCount++;
        }
    }
    
    if (processedCount === 0) {
        alert("No linked files found in selection.");
    } else {
        alert("Successfully processed " + processedCount + " linked file(s).");
    }
    
    function processLinkedFile(placedItem, hostFileName) {
        try {
            var linkedFileName = placedItem.file.name;
            
            // Remove file extension
            var dotIndex = linkedFileName.lastIndexOf(".");
            if (dotIndex > -1) {
                linkedFileName = linkedFileName.substring(0, dotIndex);
            }
            
            // Replace %20 with spaces (URL decoding)
            linkedFileName = linkedFileName.replace(/%20/g, " ");
            
            // Remove the host file prefix to isolate the material/part info
            var hostPrefix = hostFileName;
            
            // Remove "Proof", "PRIME", "ProdProof" etc. from host name to get base prefix
            hostPrefix = hostPrefix.replace(/_?(Proof|PRIME|ProdProof)$/i, "");
            
            var processedName = linkedFileName;
            
            // If the linked file starts with the host prefix, remove it
            if (linkedFileName.indexOf(hostPrefix) === 0) {
                processedName = linkedFileName.substring(hostPrefix.length);
                // Remove leading underscore if present
                if (processedName.charAt(0) === "_") {
                    processedName = processedName.substring(1);
                }
            }
            
            // Remove common keywords like CUT_, PRINT_ if they're at the beginning
            processedName = processedName.replace(/^(CUT_|PRINT_)/i, "");
            
            // Get the bounds - store them immediately
            var bounds = placedItem.geometricBounds;
            var left = bounds[0];
            var top = bounds[1]; 
            var right = bounds[2];
            var bottom = bounds[3];
            
            var placedItemWidth = Math.abs(right - left);
            var placedItemHeight = Math.abs(top - bottom);
            var widthInches = (placedItemWidth / 72) * 10;  // Multiply by 10
            var heightInches = (placedItemHeight / 72) * 10;  // Multiply by 10
            var scaledWidth = Math.round(widthInches);  // Round to whole numbers
            var scaledHeight = Math.round(heightInches);  // Round to whole numbers
            var fullName = scaledWidth + '"x' + scaledHeight + '" - ' + processedName;
            
            // Create rectangle using a completely different approach
            // Create a basic rectangle first, then directly set its properties
            
            var targetLeft = placedItem.left;
            var targetTop = placedItem.top;
            var targetWidth = placedItem.width;
            var targetHeight = placedItem.height;
            
            // Create a basic rectangle (any size)
            var whiteRect = doc.pathItems.rectangle(0, 0, -100, 100);
            
            // Now directly set its position and size properties
            whiteRect.left = targetLeft;
            whiteRect.top = targetTop;
            whiteRect.width = targetWidth;
            whiteRect.height = targetHeight;
            
            whiteRect.fillColor = createWhiteColor();
            whiteRect.filled = true;
            whiteRect.stroked = false;
            

            
            // Move rectangle behind placed item
            whiteRect.move(placedItem, ElementPlacement.PLACEAFTER);
            
            // Create text
            var textFrame = doc.textFrames.add();
            textFrame.contents = fullName;
            
            var textRange = textFrame.textRange;
            try {
                textRange.characterAttributes.textFont = app.textFonts.getByName("MyriadPro-Black");
            } catch (e) {
                try {
                    textRange.characterAttributes.textFont = app.textFonts.getByName("MyriadPro-Bold");
                } catch (e2) {
                    textRange.characterAttributes.textFont = app.textFonts[0]; 
                }
            }
            textRange.characterAttributes.size = 23;
            textRange.characterAttributes.fillColor = createBlackColor();
            textRange.paragraphAttributes.justification = Justification.CENTER;
            
            // Position text above
            var centerX = left + (placedItemWidth / 2);
            textFrame.left = centerX - (textFrame.width / 2);
            textFrame.top = top + 25;
            
            // Simple grouping - create group first
            var group = doc.groupItems.add();
            whiteRect.moveToBeginning(group);
            placedItem.moveToBeginning(group);
            textFrame.moveToBeginning(group);
            
        } catch (error) {
            alert("Error processing " + linkedFileName + ": " + error.toString());
        }
    }
    
    function extractProcessedName(linkedFileName, hostFileName) {
        var result = linkedFileName;
        var linkedDot = result.lastIndexOf(".");
        if (linkedDot > -1) {
            result = result.substring(0, linkedDot);
        }
        
        var hostNameBase = hostFileName;
        var prefixToRemove = "";

        var proofRegex = /(Proof|ProdProof)$/i;
        var match = hostNameBase.match(proofRegex);
        
        if (match) {
            var endIndex = match.index;
            prefixToRemove = hostNameBase.substring(0, endIndex);
        }

        var rawPrefixParts = prefixToRemove.split(/[_-\s]+/);
        var prefixParts = [];
        for (var k = 0; k < rawPrefixParts.length; k++) {
            if (rawPrefixParts[k].length > 0) {
                prefixParts.push(rawPrefixParts[k]);
            }
        }
        
        for (var i = 0; i < prefixParts.length; i++) {
            if (prefixParts[i].length > 0) {
                var partRegex = new RegExp("(^|[_-\\s])" + escapeRegExp(prefixParts[i]) + "([_-\\s]|$)", 'gi');
                result = result.replace(partRegex, "$1");
            }
        }

        var keywords = ["CUT", "PRINT", "PNC", "PRIME"];
        for (var j = 0; j < keywords.length; j++) {
            var keywordRegex = new RegExp("(^|[_-\\s])" + escapeRegExp(keywords[j]) + "([_-\\s]|$)", 'gi');
            result = result.replace(keywordRegex, "$1");
        }
        
        result = result.replace(/[_-\s]+/g, "_").replace(/^_|_$/g, "");
        
        return result || "Unknown";
    }

    function escapeRegExp(string) {
        return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }
    
    function createWhiteColor() {
        var color = new RGBColor();
        color.red = 255;
        color.green = 255;
        color.blue = 255;
        return color;
    }
    
    function createBlackColor() {
        var color = new RGBColor();
        color.red = 0;
        color.green = 0;
        color.blue = 0;
        return color;
    }
})();