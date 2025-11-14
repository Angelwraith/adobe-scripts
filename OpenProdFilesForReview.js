/*
@METADATA
{
  "name": "Open Production Files and organize for review",
  "description": "Open print and cut files, organize them together and tile the view to show them all",
  "version": "1.1",
  "target": "illustrator",
  "tags": ["review", "check", "file""]
}
@END_METADATA
*/

(function() {
    // Create dialog window
    var dialog = new Window("dialog", "Print/Cut File Processor");
    dialog.alignChildren = "fill";
    
    // Folder path input
    var pathGroup = dialog.add("group");
    pathGroup.orientation = "column";
    pathGroup.alignChildren = "left";
    
    pathGroup.add("statictext", undefined, "Folder Path:");
    var folderPathInput = pathGroup.add("edittext", undefined, "");
    folderPathInput.preferredSize.width = 400;
    
    // Buttons
    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "center";
    var okButton = buttonGroup.add("button", undefined, "Process Files", {name: "ok"});
    var cancelButton = buttonGroup.add("button", undefined, "Cancel", {name: "cancel"});
    
    // Show dialog
    if (dialog.show() == 1) {
        var folderPath = folderPathInput.text;
        
        if (folderPath == "") {
            alert("Please enter a folder path.");
            return;
        }
        
        var folder = new Folder(folderPath);
        if (!folder.exists) {
            alert("The specified folder does not exist.");
            return;
        }
        
        processFiles(folder);
    }
    
    function processFiles(folder) {
        try {
            // Step 1: Find and open all PRINT, PnC, and CutOnly files
            var printFiles = findFiles(folder, ["PRINT", "PnC", "CutOnly"]);
            
            if (printFiles.length == 0) {
                alert("No PRINT, PnC, or CutOnly files found.");
                return;
            }
            
            // Open all print files
            var openDocs = [];
            for (var i = 0; i < printFiles.length; i++) {
                try {
                    var doc = app.open(printFiles[i]);
                    var isCutOnly = printFiles[i].name.toUpperCase().indexOf("CUTONLY") != -1;
                    openDocs.push({
                        doc: doc,
                        file: printFiles[i],
                        baseName: getBaseName(printFiles[i].name),
                        isCutOnly: isCutOnly
                    });
                } catch (e) {
                    alert("Error opening file: " + printFiles[i].name + "\n" + e.message);
                }
            }
            
            // Step 2: Find all CUT files
            var cutFiles = findFiles(folder, ["CUT"]);
            
            if (cutFiles.length == 0) {
                alert("No CUT files found. Only PRINT files have been opened.");
                return;
            }
            
            // Step 3: Match and place CUT files into PRINT files
            var processedCount = 0;
            for (var i = 0; i < cutFiles.length; i++) {
                var cutFile = cutFiles[i];
                var cutBaseName = getBaseName(cutFile.name);
                
                // Find matching PRINT file (skip CutOnly files - they don't need a separate CUT file placed)
                var matchedDoc = null;
                for (var j = 0; j < openDocs.length; j++) {
                    if (openDocs[j].baseName == cutBaseName && !openDocs[j].isCutOnly) {
                        matchedDoc = openDocs[j].doc;
                        break;
                    }
                }
                
                if (matchedDoc != null) {
                    placeCutFile(matchedDoc, cutFile);
                    processedCount++;
                }
            }
            
            // Arrange documents in a tiled grid
            app.executeMenuCommand("tile");
            
            // Fit all documents in their windows
            for (var i = 0; i < openDocs.length; i++) {
                app.activeDocument = openDocs[i].doc;
                app.executeMenuCommand("fitall");
            }
            
            alert("Processing complete!\n" + 
                  "Opened " + openDocs.length + " PRINT files.\n" +
                  "Placed " + processedCount + " CUT files.");
                  
        } catch (e) {
            alert("Error: " + e.message);
        }
    }
    
    function findFiles(folder, keywords) {
        var foundFiles = [];
        var files = folder.getFiles();
        
        for (var i = 0; i < files.length; i++) {
            if (files[i] instanceof Folder) {
                // Recursively search subfolders
                var subFiles = findFiles(files[i], keywords);
                foundFiles = foundFiles.concat(subFiles);
            } else if (files[i] instanceof File) {
                var fileName = files[i].name.toUpperCase();
                
                // Check if file contains any of the keywords and is an AI/EPS/PDF file
                for (var k = 0; k < keywords.length; k++) {
                    if (fileName.indexOf(keywords[k].toUpperCase()) != -1 && 
                        (fileName.indexOf(".AI") != -1 || fileName.indexOf(".EPS") != -1 || fileName.indexOf(".PDF") != -1)) {
                        foundFiles.push(files[i]);
                        break;
                    }
                }
            }
        }
        
        return foundFiles;
    }
    
    function getBaseName(fileName) {
        // Remove keywords and file extension to get base name for matching
        var baseName = fileName.toUpperCase();
        baseName = baseName.replace(/PRINT/gi, "");
        baseName = baseName.replace(/PNC/gi, "");
        baseName = baseName.replace(/CUTONLY/gi, "");
        baseName = baseName.replace(/CUT/gi, "");
        baseName = baseName.replace(/\.AI$/gi, "");
        baseName = baseName.replace(/\.EPS$/gi, "");
        baseName = baseName.replace(/\.PDF$/gi, "");
        baseName = baseName.replace(/[-_\s]+/g, ""); // Remove separators
        baseName = baseName.replace(/^\s+|\s+$/g, ""); // Manual trim for ExtendScript
        
        return baseName;
    }
    
    function placeCutFile(doc, cutFile) {
        try {
            app.activeDocument = doc;
            
            // Get artboard dimensions
            var artboard = doc.artboards[0];
            var artboardRect = artboard.artboardRect;
            var artboardWidth = artboardRect[2] - artboardRect[0];
            var artboardHeight = artboardRect[1] - artboardRect[3];
            
            // Place the CUT file
            var placedItem = doc.placedItems.add();
            placedItem.file = cutFile;
            
            // Check if placed item is 1/10th the size of artboard (the Illustrator bug)
            var placedWidth = placedItem.width;
            var placedHeight = placedItem.height;
            
            var widthRatio = artboardWidth / placedWidth;
            var heightRatio = artboardHeight / placedHeight;
            
            // If either dimension is approximately 10x larger, scale up
            if ((widthRatio > 8 && widthRatio < 12) || (heightRatio > 8 && heightRatio < 12)) {
                placedItem.resize(1000, 1000); // Scale by 10x (1000%)
            }
            
            // Center the placed item on the artboard
            var itemWidth = placedItem.width;
            var itemHeight = placedItem.height;
            
            var centerX = artboardRect[0] + (artboardWidth / 2);
            var centerY = artboardRect[1] - (artboardHeight / 2);
            
            placedItem.left = centerX - (itemWidth / 2);
            placedItem.top = centerY + (itemHeight / 2);
            
            // Embed the placed item
            placedItem.embed();
            
        } catch (e) {
            alert("Error placing CUT file " + cutFile.name + ": " + e.message);
        }
    }
})();
