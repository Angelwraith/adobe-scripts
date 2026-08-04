/*@METADATA{
  "name": "Open Production Files and organize for review",
  "description": "Open print and cut files, organize them together and tile the view to show them all",
  "version": "2.0",
  "target": "illustrator",
  "tags": ["review", "check", "file"]
}@END_METADATA*/

(function() {
    // Detect OS and set Downloads path
    var downloadsPath;
    if ($.os.indexOf("Windows") !== -1) {
        // Windows path - use backslash
        var userProfile = $.getenv("USERPROFILE");
        downloadsPath = userProfile + "\\Downloads";
    } else {
        // Mac path
        downloadsPath = "~/Downloads";
    }
    
    var downloadsFolder = new Folder(downloadsPath);
    
    // Get the Filecheck Staging folder
    var stagingPath;
    if ($.os.indexOf("Windows") !== -1) {
        stagingPath = downloadsPath + "\\Filecheck Staging";
    } else {
        stagingPath = downloadsPath + "/Filecheck Staging";
    }
    
    var stagingFolder = new Folder(stagingPath);
    
    // Get list of folders in Filecheck Staging
    var subfolders = [];
    if (stagingFolder.exists) {
        var items = stagingFolder.getFiles();
        for (var i = 0; i < items.length; i++) {
            if (items[i] instanceof Folder) {
                // Ignore .stfolder folders
                if (items[i].name !== ".stfolder") {
                    subfolders.push(items[i]);
                }
            }
        }
    }

    // Load XMP library once for reading proof metadata
    try {
        if (ExternalObject.AdobeXMPScript == undefined) {
            ExternalObject.AdobeXMPScript = new ExternalObject('lib:AdobeXMPScript');
        }
        XMPMeta.registerNamespace('http://extremecolor.net/packing/', 'packing');
    } catch (e) {
        // XMP not available -- dropdown will fall back to folder names
    }

    // Build display names by reading proof file metadata from each folder's PRIME subfolder
    function getDisplayNameForFolder(folder) {
        var folderName = folder.name.replace(/%20/g, " ");
        try {
            // Look for PRIME subfolder
            var sep = ($.os.indexOf("Windows") !== -1) ? "\\" : "/";
            var primeFolder = new Folder(folder.fsName + sep + "PRIME");
            if (!primeFolder.exists) return folderName;

            // Find proof PDF/AI in PRIME folder (filename contains "Proof")
            var primeFiles = primeFolder.getFiles();
            var proofFile = null;
            for (var j = 0; j < primeFiles.length; j++) {
                if (primeFiles[j] instanceof File) {
                    var fname = primeFiles[j].name.toUpperCase();
                    if (fname.indexOf("PROOF") !== -1 && (fname.indexOf(".PDF") !== -1 || fname.indexOf(".AI") !== -1)) {
                        proofFile = primeFiles[j];
                        break;
                    }
                }
            }
            if (!proofFile) return folderName;

            // Read XMP metadata from the proof file without opening it in Illustrator
            var xmpFile = new XMPFile(proofFile.fsName, XMPConst.FILE_PDF, XMPConst.OPEN_FOR_READ);
            var xmp = xmpFile.getXMP();
            xmpFile.closeFile();

            if (!xmp.doesPropertyExist('http://extremecolor.net/packing/', 'data')) return folderName;

            var packingJSON = xmp.getProperty('http://extremecolor.net/packing/', 'data').value;
            var packingData = eval('(' + packingJSON + ')');

            // Build label: Client - Property - Project (omit property if not available)
            var parts = [];
            if (packingData.client_name) parts.push(packingData.client_name);
            if (packingData.property_name) parts.push(packingData.property_name);
            if (packingData.project_name) parts.push(packingData.project_name);

            if (parts.length > 0) return parts.join(" - ");
        } catch (e) {
            // Any error reading metadata -- fall back to folder name
        }
        return folderName;
    }

    // Create dialog window
    var dialog = new Window("dialog", "Print/Cut File Processor");
    dialog.alignChildren = "fill";
    dialog.spacing = 15;
    dialog.margins = 20;

    // Dropdown selection (if folders exist in Downloads)
    var folderDropdown = null;
    if (subfolders.length > 0) {
        var dropdownGroup = dialog.add("group");
        dropdownGroup.orientation = "column";
        dropdownGroup.alignChildren = "left";

        dropdownGroup.add("statictext", undefined, "Select folder from Downloads:");

        var dropdownRow = dropdownGroup.add("group");
        dropdownRow.orientation = "row";
        dropdownRow.spacing = 10;

        folderDropdown = dropdownRow.add("dropdownlist", undefined, []);
        folderDropdown.preferredSize.width = 400;

        for (var i = 0; i < subfolders.length; i++) {
            var displayName = getDisplayNameForFolder(subfolders[i]);
            folderDropdown.add("item", displayName);
        }
        folderDropdown.selection = 0;

        dialog.add("panel");
    }
    
    // Manual folder path input
    var pathGroup = dialog.add("group");
    pathGroup.orientation = "column";
    pathGroup.alignChildren = "left";
    
    pathGroup.add("statictext", undefined, "Or enter folder path manually:");
    var folderPathInput = pathGroup.add("edittext", undefined, "");
    folderPathInput.preferredSize.width = 400;
    
    // Buttons
    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "center";
    var openFolderButton = buttonGroup.add("button", undefined, "Open Folder");
    var okButton = buttonGroup.add("button", undefined, "Process Files", {name: "ok"});
    var cancelButton = buttonGroup.add("button", undefined, "Cancel", {name: "cancel"});
    
    // Open folder button handler
    openFolderButton.onClick = function() {
        var folderToOpen = "";
        
        // Determine which path to use
        if (folderPathInput.text != "") {
            folderToOpen = folderPathInput.text;
        } else if (folderDropdown != null && folderDropdown.selection != null) {
            var selectedIndex = folderDropdown.selection.index;
            folderToOpen = subfolders[selectedIndex].fsName;
        }
        
        if (folderToOpen == "") {
            alert("Please select a folder from the dropdown or enter a folder path first.");
            return;
        }
        
        var folder = new Folder(folderToOpen);
        if (!folder.exists) {
            alert("The specified folder does not exist.");
            return;
        }
        
        // Open folder in Windows Explorer or Mac Finder
        try {
            if ($.os.indexOf("Windows") !== -1) {
                // Windows - open in Explorer
                var batFile = new File(Folder.temp + "/openFolder.bat");
                batFile.open("w");
                batFile.writeln("@echo off");
                batFile.writeln("explorer \"" + folder.fsName + "\"");
                batFile.close();
                batFile.execute();
            } else {
                // Mac - open in Finder
                folder.execute();
            }
        } catch (e) {
            alert("Error opening folder: " + e.message);
        }
    };
    
    // Show dialog
    if (dialog.show() == 1) {
        var folderPath = "";
        
        // Determine which path to use
        if (folderPathInput.text != "") {
            // Manual path takes priority
            folderPath = folderPathInput.text;
        } else if (folderDropdown != null && folderDropdown.selection != null) {
            // Use selected dropdown folder
            var selectedIndex = folderDropdown.selection.index;
            folderPath = subfolders[selectedIndex].fsName;
        }
        
        if (folderPath == "") {
            alert("Please select a folder from the dropdown or enter a folder path.");
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

            // If nothing was found at all, alert and bail
            if (printFiles.length == 0 && cutFiles.length == 0) {
                alert("No PRINT, PnC, CutOnly, or CUT files found.");
                return;
            }

            // Check if any opened docs are non-CutOnly (i.e. PRINT/PnC that need CUT placed)
            var hasNonCutOnly = false;
            for (var i = 0; i < openDocs.length; i++) {
                if (!openDocs[i].isCutOnly) {
                    hasNonCutOnly = true;
                    break;
                }
            }

            // Step 3: Match CUT files to PRINT docs, place matched ones, open unmatched ones standalone
            for (var i = 0; i < cutFiles.length; i++) {
                var cutFile = cutFiles[i];
                var cutBaseName = getBaseName(cutFile.name);

                // Find matching PRINT/PnC file (skip CutOnly files - they don't need a separate CUT file placed)
                var matchedDoc = null;
                if (hasNonCutOnly) {
                    for (var j = 0; j < openDocs.length; j++) {
                        if (openDocs[j].baseName == cutBaseName && !openDocs[j].isCutOnly) {
                            matchedDoc = openDocs[j].doc;
                            break;
                        }
                    }
                }

                if (matchedDoc != null) {
                    // Place CUT into matching PRINT doc
                    placeCutFile(matchedDoc, cutFile);
                } else {
                    // No matching PRINT file - open the CUT file as a standalone document
                    try {
                        var cutDoc = app.open(cutFile);
                        openDocs.push({
                            doc: cutDoc,
                            file: cutFile,
                            baseName: cutBaseName,
                            isCutOnly: false
                        });
                    } catch (e) {
                        alert("Error opening CUT file: " + cutFile.name + "\n" + e.message);
                    }
                }
            }

            // Arrange documents in a tiled grid
            app.executeMenuCommand("tile");

            // Fit all documents in their windows
            for (var i = 0; i < openDocs.length; i++) {
                app.activeDocument = openDocs[i].doc;
                app.executeMenuCommand("fitall");
            }
                  
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
        baseName = baseName.replace(/_PRINT_/gi, "_");
        baseName = baseName.replace(/_PNC_/gi, "_");
        baseName = baseName.replace(/_CUTONLY_/gi, "_");
        baseName = baseName.replace(/_CUT_/gi, "_");
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