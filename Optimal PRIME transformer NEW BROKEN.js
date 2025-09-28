/*
@METADATA
{
  "name": "Optimal PRIME Transformer",
  "description": "Generate Production Files From a PRIME",
  "version": "1.4",
  "target": "illustrator",
  "tags": ["Optimal", "Prime", "processors", "scaleFactor"]
}
@END_METADATA
*/

(function() {
    // Global variables
    var doc;
    var originalDocPath; // Store original document path
    var parentSavePath; // Store the parent path for all saves
    var overwriteAll = null; // null, true, or false
    var skipAll = false;
    var scaleFactor = 1; // Add scale factor tracking
    var useFolderOrganization = false; // Track if we should organize into folders
    var materialData = {}; // Store material information for folder organization
    
    // Main execution
    try {
        // Check if document is open
        if (app.documents.length === 0) {
            alert("Please open a document before running this script.");
            return;
        }
        
        doc = app.activeDocument;
        
        // Get the scale factor for large canvas handling
        scaleFactor = doc.scaleFactor || 1; // Default to 1 if not defined
        
        // Store the original document path and calculate parent save path once
        originalDocPath = doc.fullName;
        parentSavePath = originalDocPath.parent.parent;
        
        // Check if Cut Path Separator has been run
        var cpsStatus = checkForCPSProcessing();
        
        if (!cpsStatus.hasProcessedLayers) {
            var message = "It appears the Cut Path Separator script has not been run on this document.\n\n";
            
            if (cpsStatus.foundLayers.length > 0) {
                message += "Found CPS layer(s): " + cpsStatus.foundLayers.join(", ") + "\n";
                message += "But these layers appear to be empty.\n\n";
            } else {
                message += "No CPS-created layers found.\n\n";
            }
            
            message += "It is recommended to run the Cut Path Separator script first to:\n";
            message += "â€¢ Organize cut/print data into proper layers\n";
            message += "â€¢ Standardize stroke weights\n";
            message += "â€¢ Check for duplicate paths\n";
            message += "â€¢ Enable overprint settings\n\n";
            message += "Do you want to continue anyway?";
            
            var continueProcessing = confirm(message);
            if (!continueProcessing) {
                return; // Exit the script
            }
        }
        
        // Check for PRIME in filename
        if (!checkForPrime()) {
            return;
        }
        
        // Check artboard names
        if (!checkArtboardNames()) {
            return;
        }
        
        // Check for illegal characters in artboard names
        if (!checkArtboardCharacters()) {
            return;
        }
        
		// Check if we should use folder organization
        if (doc.artboards.length > 1) {
            // Validate material naming and get summary in one step
            var materialValidation = validateMaterialNaming();
            
            if (!materialValidation.isValid) {
                var errorMessage = "âš ï¸ ARTBOARD NAMING ISSUES DETECTED:\n\n";
                errorMessage += materialValidation.errors.join("\n") + "\n\n";
                errorMessage += "Correct format: 'Material_PtX' where:\n";
                errorMessage += "â€¢ Material can contain letters, numbers, spaces, and underscores\n";
                errorMessage += "â€¢ _PtX is optional for single parts (X can be any number)\n";
                errorMessage += "â€¢ Part numbers must be unique for each material\n\n";
                errorMessage += "Please fix the artboard names before running the script again.";
                
                alert(errorMessage);
                return;
            } else {
                // Show folder organization dialog
                var folderChoice = showFolderOrganizationDialog(materialValidation.materialData);
                
                if (folderChoice === null) {
                    return; // User cancelled
                } else if (folderChoice === true) {
                    useFolderOrganization = true;
                    materialData = materialValidation.materialData;
                } else {
                    useFolderOrganization = false;
                    materialData = {};
                }
            }
        }
		
        // Process each artboard
        processAllArtboards();
        
        alert("Processing complete!" + (scaleFactor !== 1 ? "\nLarge Canvas scaling factor (" + scaleFactor + ") has been applied." : ""));
        
    } catch (e) {
        alert("Error occurred: " + e.toString() + "\nLine: " + e.line);
    }
	
	function showFolderOrganizationDialog(materialValidation) {
        var dialog = new Window("dialog", "Folder Organization");
        dialog.orientation = "column";
        dialog.alignChildren = "fill";
        dialog.spacing = 15;
        dialog.margins = 20;
        dialog.preferredSize.width = 400;
        
        // Header
        var headerText = dialog.add("statictext", undefined, "This document has " + doc.artboards.length + " artboards.");
        headerText.graphics.font = ScriptUI.newFont("dialog", "bold", 12);
        
        // Materials summary
        var summaryGroup = dialog.add("group");
        summaryGroup.orientation = "column";
        summaryGroup.alignChildren = "fill";
        
        summaryGroup.add("statictext", undefined, "Materials detected:");
        
        var materialsList = summaryGroup.add("edittext", undefined, "", {multiline: true, readonly: true});
        materialsList.preferredSize.height = 120;
        
		var materialsText = "";
				for (var materialName in materialValidation) {
					var materialInfo = materialValidation[materialName];
					var count = materialInfo.artboards.length;
					var frontBackCount = 0;
					
					// Count front/back pairs
					if (materialInfo.frontBackPairs) {
						for (var pairKey in materialInfo.frontBackPairs) {
							var pair = materialInfo.frontBackPairs[pairKey];
							if (pair.FRONT || pair.BACK) {
								frontBackCount++;
							}
						}
					}
					
					materialsText += "â€¢ " + materialName + ": " + count + " part" + (count > 1 ? "s" : "");
					if (frontBackCount > 0) {
						materialsText += " (" + frontBackCount + " with front/back)";
					}
					materialsText += "\n";
					
					if (count > 10 && !materialUsesCutContour(materialName)) {
						materialsText += "  â†' Will create PRINT and CUT subfolders\n";
					}
				}
        materialsList.text = materialsText;
        
        // Instructions
        var instructText = dialog.add("statictext", undefined, "Choose how to organize the output files:");
        instructText.graphics.font = ScriptUI.newFont("dialog", "bold", 11);
        
        // Button group
        var buttonGroup = dialog.add("group");
        buttonGroup.alignment = "center";
        buttonGroup.spacing = 10;
        
        var foldersButton = buttonGroup.add("button", undefined, "Create Folders");
        foldersButton.preferredSize.width = 100;
        
        var rootButton = buttonGroup.add("button", undefined, "Root Directory");
        rootButton.preferredSize.width = 100;
        
        var cancelButton = buttonGroup.add("button", undefined, "Cancel");
        cancelButton.preferredSize.width = 80;
        
        // Button handlers
        var result = null;
        
        foldersButton.onClick = function() {
            result = true;
            dialog.close();
        };
        
        rootButton.onClick = function() {
            result = false;
            dialog.close();
        };
        
        cancelButton.onClick = function() {
            result = null;
            dialog.close();
        };
        
        // Show dialog
        dialog.show();
        
        return result;
    }
    
    // Check if Cut Path Separator has been run by looking for CPS-created layers with content
    function checkForCPSProcessing() {
        var cpsLayerNames = ["CutThrough2-Outside", "CutThrough1-Inside", "CutThrough-Knifecut", "CutContour", "Spot1"];
        var foundLayers = [];
        var layersWithContent = [];
        
        // Check each CPS layer
        for (var i = 0; i < doc.layers.length; i++) {
            var layer = doc.layers[i];
            var layerName = layer.name;
            
            // Check if this is a CPS-created layer
            for (var j = 0; j < cpsLayerNames.length; j++) {
                if (layerName === cpsLayerNames[j]) {
                    foundLayers.push(layerName);
                    
                    // Check if layer has content using the existing layerHasContent function
                    if (layerHasContent(layer)) {
                        layersWithContent.push(layerName);
                    }
                    break;
                }
            }
        }
        
        return {
            foundLayers: foundLayers,
            layersWithContent: layersWithContent,
            hasProcessedLayers: layersWithContent.length > 0
        };
    }
    
    // Detect front/back designation in artboard names
    function detectFrontBackDesignation(artboardName) {
        var lowerName = artboardName.toLowerCase();
        
// Common front variations - improved patterns for your format
        var frontPatterns = [
            /_front$/i,      // _Front at end
            /_frt$/i,        // _Frt at end  
            /_f$/i,          // _F at end
            /\bfront\b/i,    // Front as word
            /\bfrt\b/i       // Frt as word
        ];
        
        // Common back variations - improved patterns for your format
        var backPatterns = [
            /_back$/i,       // _Back at end
            /_bk$/i,         // _Bk at end
            /_b$/i,          // _B at end
            /\bback\b/i,     // Back as word
            /\bbk\b/i        // Bk as word
        ];
        
        // Check for front patterns
        for (var i = 0; i < frontPatterns.length; i++) {
            if (frontPatterns[i].test(artboardName)) {
                return "FRONT";
            }
        }
        
        // Check for back patterns
        for (var i = 0; i < backPatterns.length; i++) {
            if (backPatterns[i].test(artboardName)) {
                return "BACK";
            }
        }
        
        return null; // No front/back designation found
    }
    
// Strip front/back designation from artboard name to get base name
    function stripFrontBackDesignation(artboardName) {
        // Remove common front/back suffixes - more specific patterns first
        var cleanName = artboardName
            .replace(/_front$/i, "")    // Remove _Front at end
            .replace(/_frt$/i, "")      // Remove _Frt at end  
            .replace(/_f$/i, "")        // Remove _F at end
            .replace(/_back$/i, "")     // Remove _Back at end
            .replace(/_bk$/i, "")       // Remove _Bk at end
            .replace(/_b$/i, "")        // Remove _B at end
            .replace(/\s+/g, "_");      // Replace spaces with underscores
        
        return cleanName;
    }
    
// Validate material naming format for folder organization
    function validateMaterialNaming() {
        var materialData = {};
        var errors = [];
        var isValid = true;
        
        for (var i = 0; i < doc.artboards.length; i++) {
            var artboardName = doc.artboards[i].name;
            if (typeof artboardName === 'string') {
                artboardName = artboardName.replace(/^\s+|\s+$/g, ''); // Manual trim for compatibility
            }
            
            // Parse material name and part number - ExtendScript compatible approach
            var materialName, partNumber, frontBackDesignation;
            
            // Handle the front/back stripping using indexOf instead of endsWith
            var workingName = artboardName;
            var lowerName = workingName.toLowerCase();
            
            if (lowerName.indexOf('_front') === lowerName.length - 6) {
                frontBackDesignation = 'FRONT';
                workingName = workingName.substring(0, workingName.length - 6); // Remove '_Front'
            } else if (lowerName.indexOf('_back') === lowerName.length - 5) {
                frontBackDesignation = 'BACK';
                workingName = workingName.substring(0, workingName.length - 5); // Remove '_Back'
            } else if (lowerName.indexOf('_frt') === lowerName.length - 4) {
                frontBackDesignation = 'FRONT';
                workingName = workingName.substring(0, workingName.length - 4); // Remove '_Frt'
            } else if (lowerName.indexOf('_bk') === lowerName.length - 3) {
                frontBackDesignation = 'BACK';
                workingName = workingName.substring(0, workingName.length - 3); // Remove '_Bk'
            } else {
                frontBackDesignation = null;
            }
            
            // Set baseArtboardName for use elsewhere in the code
            var baseArtboardName = workingName;
            
            // Check the cleaned artboard name for illegal characters
            var allowedPattern = /^[a-zA-Z0-9\s_]+(_Pt\d+)?$/i;
            if (!allowedPattern.test(baseArtboardName)) {
                // Create a more robust illegal character detection
                var illegalChars = "";
                var charAllowedPattern = /[a-zA-Z0-9\s_]/;
                
                for (var charIndex = 0; charIndex < baseArtboardName.length; charIndex++) {
                    var currentChar = baseArtboardName.charAt(charIndex);
                    if (!charAllowedPattern.test(currentChar)) {
                        illegalChars += currentChar;
                    }
                }
                
                // Handle _Pt part separately
                if (baseArtboardName.indexOf('_Pt') !== -1) {
                    var parts = baseArtboardName.split('_Pt');
                    if (parts.length === 2 && !/^\d+$/.test(parts[1])) {
                        errors.push("Artboard " + (i + 1) + " '" + artboardName + "': Part number after '_Pt' must be numeric");
                        isValid = false;
                        continue;
                    }
                }
                
                if (illegalChars.length > 0) {
                    var uniqueChars = [];
                    for (var c = 0; c < illegalChars.length; c++) {
                        var found = false;
                        for (var u = 0; u < uniqueChars.length; u++) {
                            if (uniqueChars[u] === illegalChars.charAt(c)) {
                                found = true;
                                break;
                            }
                        }
                        if (!found) {
                            uniqueChars.push(illegalChars.charAt(c));
                        }
                    }
                    errors.push("Artboard " + (i + 1) + " '" + artboardName + "': Contains illegal characters: " + uniqueChars.join(" "));
                    isValid = false;
                    continue;
                }
            }
            
            // Now parse the cleaned name for material and part number
            var ptIndex = workingName.toLowerCase().indexOf('_pt');
            
            if (ptIndex !== -1) {
                materialName = baseArtboardName.substring(0, ptIndex);
                partNumber = baseArtboardName.substring(ptIndex + 3);
                
                // Validate part number is numeric
                if (!/^\d+$/.test(partNumber)) {
                    errors.push("Artboard " + (i + 1) + " '" + artboardName + "': Part number must be numeric");
                    isValid = false;
                    continue;
                }
            } else {
                materialName = baseArtboardName;
                partNumber = null; // No part number means single part
            }
            
            // Initialize material data if first occurrence
            if (!materialData[materialName]) {
                materialData[materialName] = {
                    artboards: [],
                    partNumbers: [],
                    originalCase: materialName, // Store first occurrence case
                    frontBackPairs: {} // Track front/back pairs
                };
            }
            
            // Check for duplicate part numbers within the same material
            if (partNumber !== null) {
                var partNumbers = materialData[materialName].partNumbers;
                var isDuplicate = false;
                for (var p = 0; p < partNumbers.length; p++) {
                    if (partNumbers[p] === partNumber) {
                        isDuplicate = true;
                        break;
                    }
                }
                if (isDuplicate) {
                    errors.push("Material '" + materialName + "': Duplicate part number " + partNumber);
                    isValid = false;
                    continue;
                }
                materialData[materialName].partNumbers.push(partNumber);
            }
            
            // Add artboard to material data
            var artboardData = {
                index: i,
                name: artboardName,
                partNumber: partNumber,
                frontBack: frontBackDesignation,
                baseName: baseArtboardName
            };
            
            materialData[materialName].artboards.push(artboardData);
            
            // Track front/back pairs for validation
            if (frontBackDesignation) {
                var pairKey = partNumber || "single";
                if (!materialData[materialName].frontBackPairs[pairKey]) {
                    materialData[materialName].frontBackPairs[pairKey] = {};
                }
                
                if (materialData[materialName].frontBackPairs[pairKey][frontBackDesignation]) {
                    errors.push("Material '" + materialName + "': Duplicate " + frontBackDesignation + " for part " + (partNumber || "single"));
                    isValid = false;
                    continue;
                }
                
                materialData[materialName].frontBackPairs[pairKey][frontBackDesignation] = artboardData;
            }
        }
        
        return {
            isValid: isValid,
            errors: errors,
            materialData: materialData
        };
    }
	
	// Check if a specific material uses CutContour by examining its artboards
	function checkMaterialUsesCutContour(materialName) {
		// Get the list of artboard indices for this material
		var materialInfo = materialData[materialName];
		if (!materialInfo || !materialInfo.artboards) {
			return false;
		}
		
		// Check each artboard for this material to see if any have CutContour content
		for (var i = 0; i < materialInfo.artboards.length; i++) {
			var artboardIndex = materialInfo.artboards[i].index;
			
			// Set active artboard to check its content
			var originalActiveIndex = doc.artboards.getActiveArtboardIndex();
			doc.artboards.setActiveArtboardIndex(artboardIndex);
			
			// Select all objects on this artboard
			doc.selectObjectsOnActiveArtboard();
			
			// Check if any selected objects are on CutContour layers
			for (var j = 0; j < doc.selection.length; j++) {
				var item = doc.selection[j];
				if (item.layer && (item.layer.name === "CutContour" || item.layer.name === "CutContour_PRF")) {
					// Found CutContour content for this material
					doc.selection = null; // Clear selection
					doc.artboards.setActiveArtboardIndex(originalActiveIndex); // Restore original artboard
					return true;
				}
			}
			
			// Clear selection and restore original artboard
			doc.selection = null;
			doc.artboards.setActiveArtboardIndex(originalActiveIndex);
		}
		
		return false; // No CutContour content found for this material
	}
 
    // Check if a material uses CutContour (which generates PnC files, not separate PRINT/CUT)
    function materialUsesCutContour(materialName) {
        // Find layers that might indicate CutContour usage
        for (var i = 0; i < doc.layers.length; i++) {
            var layer = doc.layers[i];
            if ((layer.name === "CutContour" || layer.name === "CutContour_PRF") && layerHasContent(layer)) {
                return true;
            }
        }
        return false;
    }
    
    // Check if filename contains PRIME
    function checkForPrime() {
        var fileName = doc.name;
        if (fileName.indexOf("PRIME") === -1) {
            alert("File name must contain 'PRIME'. Please save the file as a PRIME file before continuing.");
            return false;
        }
        return true;
    }
    
    // Check if artboards are properly named
    function checkArtboardNames() {
        var hasUnnamedArtboards = false;
        var unnamedList = [];
        
        for (var i = 0; i < doc.artboards.length; i++) {
            var artboardName = doc.artboards[i].name;
            var artboardNameLower = artboardName.toLowerCase();
            // Check if name is just a number, contains "artboard" (case-insensitive), or is "print" (case-insensitive)
            if (/^\d+$/.test(artboardName) || /artboard/i.test(artboardName) || artboardNameLower === "print") {
                hasUnnamedArtboards = true;
                unnamedList.push("Artboard " + (i + 1) + ": '" + artboardName + "'");
            }
        }
        
        if (hasUnnamedArtboards) {
            var message = "The following artboards are not properly named:\n" + 
                         unnamedList.join("\n") + 
                         "\n\nDo you want to continue anyway?";
            var firstConfirm = confirm(message);
            
            if (firstConfirm) {
                // Second warning
                var secondConfirm = confirm("Are you REALLY sure?\n\nIt is HIGHLY recommended that you cancel and correct file names first.\n\n'Whatever! I Do What I Want!'\n\nYes to do what you want. No to stop.");
                
                if (secondConfirm) {
                    // Third warning with random number verification
                    var randomNumber = Math.floor(Math.random() * 900000000) + 100000000; // Generate 9-digit random number
                    var userInput = prompt("Samir! You're breaking the car!\n\nVerification Number: " + randomNumber + "\n\nEnter the number above to continue or click Cancel to fix the artboard names:");
                    
                    if (userInput === null) {
                        return false; // User clicked Cancel
                    }
                    
                    if (userInput === randomNumber.toString()) {
                        return true; // User entered correct number
                    } else {
                        alert("Incorrect number entered.\n\nPlease fix your artboard names.");
                        return false;
                    }
                }
                return false;
            }
            return false;
        }
        return true;
    }
    
    // Check if artboards have illegal characters for file systems
    function checkArtboardCharacters() {
        var hasIllegalChars = false;
        var problemList = [];
        // Using string method instead of regex to avoid escape issues
        var illegalCharsString = '<>:"/\\|?*';
        
        for (var i = 0; i < doc.artboards.length; i++) {
            var artboardName = doc.artboards[i].name;
            var foundChars = [];
            
            // Check each illegal character
            for (var j = 0; j < illegalCharsString.length; j++) {
                if (artboardName.indexOf(illegalCharsString[j]) !== -1) {
                    if (foundChars.indexOf(illegalCharsString[j]) === -1) {
                        foundChars.push(illegalCharsString[j]);
                    }
                }
            }
            
            if (foundChars.length > 0) {
                hasIllegalChars = true;
                problemList.push("Artboard " + (i + 1) + " '" + artboardName + "' contains: " + foundChars.join(" "));
            }
        }
        
        if (hasIllegalChars) {
            alert("The following artboards have illegal characters for file names:\n\n" + 
                  problemList.join("\n") + 
                  "\n\nPlease rename these artboards to remove illegal characters.");
            return false;
        }
        return true;
    }
    
    // Process all artboards
    function processAllArtboards() {
        var totalArtboards = doc.artboards.length;
        
        for (var i = 0; i < totalArtboards; i++) {
            try {
                // Set active artboard
                doc.artboards.setActiveArtboardIndex(i);
                var artboardName = doc.artboards[i].name;
                
                // Isolate current artboard  
                var hasContent = isolateArtboard(i);
                
                // Skip if artboard is empty
                if (!hasContent) {
                    continue;
                }
                
                // Force completion of isolation
                app.redraw();
                
                // Save isolated state to temp file for duplication
                var tempPath = Folder.temp + "/isolated_" + Date.now() + ".ai";
                var tempFile = new File(tempPath);
                var aiOptions = new IllustratorSaveOptions();
                aiOptions.compatibility = Compatibility.ILLUSTRATOR24;
                doc.saveAs(tempFile, aiOptions);
                
                // Check what type of file to create
                // Check if CutContour or CutContour_PRF layers have content on this artboard
                var cutContourInfo = getCutContourInfo();
                
                if (cutContourInfo.hasCutContour) {
                    if (cutContourInfo.hasPrintData) {
                        createPnCFileFromTemp(tempFile, artboardName);
                    } else {
                        createCutOnlyFileFromTemp(tempFile, artboardName);
                    }
                } else {
                    createPrintFileFromTemp(tempFile, artboardName);
                    createCutFileFromTemp(tempFile, artboardName);
                }
                
                // Delete temp file
                tempFile.remove();
                
                // Reopen original document
                doc = app.open(originalDocPath);
                
                // Restore scale factor reference after reopening
                scaleFactor = doc.scaleFactor || 1;
                
            } catch (e) {
                alert("Error processing artboard " + (i + 1) + " (" + artboardName + "): " + 
                      e.toString() + "\nStopping process.");
                throw e;
            }
        }
    }
    
    // Isolate current artboard
    function isolateArtboard(artboardIndex) {
        // Store the current artboard's bounds for reference
        var currentArtboard = doc.artboards[artboardIndex];
        var artboardBounds = currentArtboard.artboardRect;
        
        // Select all on active artboard
        doc.artboards.setActiveArtboardIndex(artboardIndex);
        doc.selectObjectsOnActiveArtboard();
        
        // If nothing selected on this artboard, skip it
        if (doc.selection.length === 0) {
            return false; // Return false to indicate empty artboard
        }
        
        // Select inverse and delete
        app.executeMenuCommand("Inverse menu item");
        
        // Delete inverse selection
        if (doc.selection.length > 0) {
            for (var j = doc.selection.length - 1; j >= 0; j--) {
                doc.selection[j].remove();
            }
        }
        
        // Clear selection
        doc.selection = null;
        
        // Remove all other artboards but keep the current one
        var newIndex = 0;
        for (var k = doc.artboards.length - 1; k >= 0; k--) {
            if (k !== artboardIndex) {
                doc.artboards.remove(k);
                if (k < artboardIndex) {
                    newIndex = artboardIndex - 1;
                } else {
                    newIndex = artboardIndex;
                }
                artboardIndex = newIndex;
            }
        }
        
        doc.artboards.setActiveArtboardIndex(0);
        doc.selection = null;
        
        return true;
    }
    
	// Get information about CutContour layers and print data
	function getCutContourInfo() {
		var cutContourLayer = null;
		var cutContourPrfLayer = null;
		var printLayers = []; // Changed to array to handle multiple print layer types
		
		// Find relevant layers
		for (var i = 0; i < doc.layers.length; i++) {
			var layerName = doc.layers[i].name;
			var layerNameLower = layerName.toLowerCase();
			
			if (layerName === "CutContour") {
				cutContourLayer = doc.layers[i];
			} else if (layerName === "CutContour_PRF") {
				cutContourPrfLayer = doc.layers[i];
			} else if (layerNameLower === "print" || layerNameLower === "spot1") {
				// Now recognizes both "print" and "spot1" as print layers
				printLayers.push(doc.layers[i]);
			}
		}
		
		// Check if any CutContour layers have content
		var hasCutContour = false;
		if (cutContourLayer && layerHasContent(cutContourLayer)) {
			hasCutContour = true;
		}
		if (cutContourPrfLayer && layerHasContent(cutContourPrfLayer)) {
			hasCutContour = true;
		}
		
		// Check if any print layers have content
		var hasPrintData = false;
		for (var j = 0; j < printLayers.length; j++) {
			if (layerHasContent(printLayers[j])) {
				hasPrintData = true;
				break; // Found at least one print layer with content
			}
		}
		
		return {
			hasCutContour: hasCutContour,
			hasPrintData: hasPrintData
		};
    }
    
    // Create PRINT file from temp file
    function createPrintFileFromTemp(tempFile, artboardName) {
        // PRESERVE CONTEXT before opening temp file
        var savedMaterialData = materialData;
        var savedUseFolderOrganization = useFolderOrganization;
        var savedParentSavePath = parentSavePath;
        var savedOverwriteAll = overwriteAll;
        var savedSkipAll = skipAll;
        
        // Open the temp file
        var tempDoc = app.open(tempFile);
        
        var hasAnyContent = false;
        
        // Delete layers we don't want for PRINT file
        for (var i = tempDoc.layers.length - 1; i >= 0; i--) {
            var layer = tempDoc.layers[i];
            var layerName = layer.name.toLowerCase();
            
            // Delete cut layers (including CutContour and CutContour_PRF)
            if (layerName.indexOf("cut") !== -1) {
                layer.remove();
                continue;
            }
            
            // Keep only print/reg/spot1 that have content
            if (layerName === "print" || layerName === "reg" || layerName === "spot1") {
                if (!layerHasContent(layer)) {
                    layer.remove();
                } else {
                    hasAnyContent = true;
                }
            } else {
                layer.remove();
            }
        }
        
        // Check if document has any content left
        if (!hasAnyContent || tempDoc.layers.length === 0) {
            // Skip this file - no content
            tempDoc.close(SaveOptions.DONOTSAVECHANGES);
            return;
        }
        
        // Save as PDF with preserved context
        var fileName = generateFileName("PRINT", artboardName);
        saveFileAsPDFWithContext(tempDoc, fileName, artboardName, 
                                savedMaterialData, savedUseFolderOrganization, 
                                savedParentSavePath, savedOverwriteAll, savedSkipAll);
        
        // Close without saving
        tempDoc.close(SaveOptions.DONOTSAVECHANGES);
    }
    
    // Create CUT file from temp file
    function createCutFileFromTemp(tempFile, artboardName) {
        // PRESERVE CONTEXT before opening temp file
        var savedMaterialData = materialData;
        var savedUseFolderOrganization = useFolderOrganization;
        var savedParentSavePath = parentSavePath;
        var savedOverwriteAll = overwriteAll;
        var savedSkipAll = skipAll;
        
        // Open the temp file
        var tempDoc = app.open(tempFile);
        
        var hasKnifeCut = false;
        var otherCuts = [];
        var hasAnyContent = false;
        
        // Delete and check layers
        for (var i = tempDoc.layers.length - 1; i >= 0; i--) {
            var layer = tempDoc.layers[i];
            var layerName = layer.name;
            var layerNameLower = layerName.toLowerCase();
            
            // Delete PRINT and Spot1
            if (layerNameLower === "print" || layerNameLower === "spot1") {
                layer.remove();
                continue;
            }
            
            // Delete CutContour and CutContour_PRF layers (they should go to PnC or separate files)
            if (layerName === "CutContour" || layerName === "CutContour_PRF") {
                layer.remove();
                continue;
            }
            
            // Keep REG if it has content
            if (layerNameLower === "reg") {
                if (!layerHasContent(layer)) {
                    layer.remove();
                } else {
                    hasAnyContent = true;
                }
                continue;
            }
            
            // Process other cut layers (excluding CutContour types)
            if (layerNameLower.indexOf("cut") !== -1) {
                if (!layerHasContent(layer)) {
                    layer.remove();
                } else {
                    hasAnyContent = true;
                    if (layerName === "CutThrough3-Knifecut") {
                        hasKnifeCut = true;
                    } else {
                        otherCuts.push(layerName);
                    }
                }
            } else {
                layer.remove();
            }
        }
        
        // Check if document has any content left
        if (!hasAnyContent || tempDoc.layers.length === 0) {
            // Skip this file - no content
            tempDoc.close(SaveOptions.DONOTSAVECHANGES);
            return;
        }
        
        // Warn if knife cut is with other cuts
        if (hasKnifeCut && otherCuts.length > 0) {
            alert("Warning: Knife cut (CutThrough3-Knifecut) is on the same artboard as other cut types:\n" + 
                  otherCuts.join(", "));
        }
        
        // Save as PDF with preserved context
        var fileName = generateFileName("CUT", artboardName);
        saveFileAsPDFWithContext(tempDoc, fileName, artboardName, 
                                savedMaterialData, savedUseFolderOrganization, 
                                savedParentSavePath, savedOverwriteAll, savedSkipAll);
        
        // Close without saving
        tempDoc.close(SaveOptions.DONOTSAVECHANGES);
    }
    
    // Create PnC file from temp file (when CutContour/CutContour_PRF + Print data present)
    function createPnCFileFromTemp(tempFile, artboardName) {
        // PRESERVE CONTEXT before opening temp file
        var savedMaterialData = materialData;
        var savedUseFolderOrganization = useFolderOrganization;
        var savedParentSavePath = parentSavePath;
        var savedOverwriteAll = overwriteAll;
        var savedSkipAll = skipAll;
        
        // Open the temp file
        var tempDoc = app.open(tempFile);
        
        var hasAnyContent = false;
        
        // Delete layers we don't want
        for (var i = tempDoc.layers.length - 1; i >= 0; i--) {
            var layer = tempDoc.layers[i];
            var layerName = layer.name;
            var layerNameLower = layerName.toLowerCase();
            
            // Keep only print/spot1/cutcontour/cutcontour_prf that have content
            if (layerNameLower === "print" || layerNameLower === "spot1" || 
                layerName === "CutContour" || layerName === "CutContour_PRF") {
                if (!layerHasContent(layer)) {
                    layer.remove();
                } else {
                    hasAnyContent = true;
                }
            } else {
                layer.remove();
            }
        }
        
        // Check if document has any content left
        if (!hasAnyContent || tempDoc.layers.length === 0) {
            // Skip this file - no content
            tempDoc.close(SaveOptions.DONOTSAVECHANGES);
            return;
        }
        
        // Save as PDF with preserved context
        var fileName = generateFileName("PnC", artboardName);
        saveFileAsPDFWithContext(tempDoc, fileName, artboardName, 
                                savedMaterialData, savedUseFolderOrganization, 
                                savedParentSavePath, savedOverwriteAll, savedSkipAll);
        
        // Close without saving
        tempDoc.close(SaveOptions.DONOTSAVECHANGES);
    }
    
    // Create CutOnly file from temp file (when CutContour/CutContour_PRF present but no Print data)
    function createCutOnlyFileFromTemp(tempFile, artboardName) {
        // PRESERVE CONTEXT before opening temp file
        var savedMaterialData = materialData;
        var savedUseFolderOrganization = useFolderOrganization;
        var savedParentSavePath = parentSavePath;
        var savedOverwriteAll = overwriteAll;
        var savedSkipAll = skipAll;
        
        // Open the temp file
        var tempDoc = app.open(tempFile);
        
        var hasAnyContent = false;
        
        // Delete layers we don't want
        for (var i = tempDoc.layers.length - 1; i >= 0; i--) {
            var layer = tempDoc.layers[i];
            var layerName = layer.name;
            var layerNameLower = layerName.toLowerCase();
            
            // Keep only cutcontour/cutcontour_prf/reg that have content
            if (layerName === "CutContour" || layerName === "CutContour_PRF" || layerNameLower === "reg") {
                if (!layerHasContent(layer)) {
                    layer.remove();
                } else {
                    hasAnyContent = true;
                }
            } else {
                layer.remove();
            }
        }
        
        // Check if document has any content left
        if (!hasAnyContent || tempDoc.layers.length === 0) {
            // Skip this file - no content
            tempDoc.close(SaveOptions.DONOTSAVECHANGES);
            return;
        }
        
        // Save as PDF with preserved context
        var fileName = generateFileName("CutOnly", artboardName);
        saveFileAsPDFWithContext(tempDoc, fileName, artboardName, 
                                savedMaterialData, savedUseFolderOrganization, 
                                savedParentSavePath, savedOverwriteAll, savedSkipAll);
        
        // Close without saving
        tempDoc.close(SaveOptions.DONOTSAVECHANGES);
    }
    
    // Check if layer has content
    function layerHasContent(layer) {
        if (layer.pageItems.length > 0) {
            return true;
        }
        
        // Check sublayers
        for (var i = 0; i < layer.layers.length; i++) {
            if (layerHasContent(layer.layers[i])) {
                return true;
            }
        }
        
        return false;
    }
    
    // New saveFileAsPDFWithContext function that accepts preserved context
    function saveFileAsPDFWithContext(tempDoc, fileName, artboardName, 
                                     contextMaterialData, contextUseFolderOrganization, 
                                     contextParentSavePath, contextOverwriteAll, contextSkipAll) {
        var saveFile;
        
        // Use the preserved context instead of global variables
        if (contextUseFolderOrganization && artboardName) {
            var materialName = getMaterialNameFromArtboardName(artboardName);
            var materialFolder = createMaterialFolderWithContext(materialName, contextParentSavePath);
            
            if (!materialFolder) {
                // Folder creation failed, ask user what to do
                var fallbackChoice = confirm("Failed to create folder for material '" + materialName + "'.\n\n" +
                                           "Click OK to save in the main directory, or Cancel to stop processing.");
                if (!fallbackChoice) {
                    return false; // Stop processing
                }
                // Continue with main directory
                saveFile = new File(contextParentSavePath + "/" + fileName);
            } else {
                // Check if THIS SPECIFIC material needs subfolders using preserved context
                var materialInfo = contextMaterialData[materialName];
                var needsSubfolders = false;
                
                // Only create subfolders if this material has >10 parts
                // (We can't check CutContour from temp doc context, so simplified logic)
                if (materialInfo && materialInfo.artboards.length > 10) {
                    needsSubfolders = true;
                }
                
                if (needsSubfolders) {
                    // Create subfolders for PRINT and CUT files only
                    if (fileName.indexOf("_PRINT_") !== -1) {
                        var printFolder = createSubFolder(materialFolder, "PRINT");
                        saveFile = new File((printFolder || materialFolder) + "/" + fileName);
                    } else if (fileName.indexOf("_CUT_") !== -1) {
                        var cutFolder = createSubFolder(materialFolder, "CUT");
                        saveFile = new File((cutFolder || materialFolder) + "/" + fileName);
                    } else {
                        // PnC or CutOnly files go in material root folder
                        saveFile = new File(materialFolder + "/" + fileName);
                    }
                } else {
                    // This material has ≤10 parts, no subfolders needed
                    saveFile = new File(materialFolder + "/" + fileName);
                }
            }
        } else {
            // Use the preserved parent save path
            saveFile = new File(contextParentSavePath + "/" + fileName);
        }
        
        // Check if file exists using preserved overwrite settings
        if (saveFile.exists) {
            var overwrite = false;
            
            if (contextOverwriteAll === null) {
                var dialog = confirm("File '" + fileName + "' already exists.\nDo you want to overwrite it?\n\n" +
                                   "Click 'Yes' to overwrite, 'No' to skip.");
                
                if (dialog) {
                    overwrite = true;
                    // Ask if should apply to all
                    if (confirm("Apply this decision to all existing files?")) {
                        // Update global variable
                        overwriteAll = true;
                    }
                } else {
                    if (confirm("Skip all existing files?")) {
                        // Update global variable
                        skipAll = true;
                    }
                }
            } else {
                overwrite = contextOverwriteAll;
            }
            
            if (!overwrite || contextSkipAll) {
                return; // Skip this file
            }
        }
        
        // Set up PDF save options with Large Canvas compatibility
        var pdfOptions = new PDFSaveOptions();
        pdfOptions.compatibility = PDFCompatibility.ACROBAT7; // PDF 1.6 for Large Canvas support
        pdfOptions.generateThumbnails = true;
        pdfOptions.preserveEditability = true;
        pdfOptions.preset = "[Illustrator Default]";
        pdfOptions.viewAfterSaving = false;
        
        // Save as PDF directly
        tempDoc.saveAs(saveFile, pdfOptions);
        
        // Force completion
        app.redraw();
    }
    
    // Helper function to create material folder with specific parent path
    function createMaterialFolderWithContext(materialName, contextParentSavePath) {
        try {
            var materialFolder = new Folder(contextParentSavePath + "/" + materialName);
            if (!materialFolder.exists) {
                materialFolder.create();
            }
            return materialFolder.fsName;
        } catch (e) {
            return null; // Folder creation failed
        }
    }
    
    // Save document as PDF with Large Canvas scaling fix and folder organization (LEGACY - kept for compatibility)
	function saveFileAsPDF(tempDoc, fileName, artboardName) {
		var saveFile;
		
		// Determine save path based on folder organization setting
		if (useFolderOrganization && artboardName) {
			var materialName = getMaterialNameFromArtboardName(artboardName);
			var materialFolder = createMaterialFolder(materialName);
			
			if (!materialFolder) {
				// Folder creation failed, ask user what to do
				var fallbackChoice = confirm("Failed to create folder for material '" + materialName + "'.\n\n" +
										   "Click OK to save in the main directory, or Cancel to stop processing.");
				if (!fallbackChoice) {
					return false; // Stop processing
				}
				// Continue with main directory
				saveFile = new File(parentSavePath + "/" + fileName);
			} else {
				// Check if THIS SPECIFIC material needs subfolders 
				// (>10 artboards for this material AND not using CutContour)
				var materialInfo = materialData[materialName];
				var needsSubfolders = false;
				
				// Only create subfolders if this material has >10 parts AND doesn't use CutContour
				if (materialInfo && materialInfo.artboards.length > 10) {
					// Check if this specific material uses CutContour by checking its artboards
					var materialUsesCutContourFlag = checkMaterialUsesCutContour(materialName);
					if (!materialUsesCutContourFlag) {
						needsSubfolders = true;
					}
				}
				
				if (needsSubfolders) {
					// Create subfolders for PRINT and CUT files only
					if (fileName.indexOf("_PRINT_") !== -1) {
						var printFolder = createSubFolder(materialFolder, "PRINT");
						saveFile = new File((printFolder || materialFolder) + "/" + fileName);
					} else if (fileName.indexOf("_CUT_") !== -1) {
						var cutFolder = createSubFolder(materialFolder, "CUT");
						saveFile = new File((cutFolder || materialFolder) + "/" + fileName);
					} else {
						// PnC or CutOnly files go in material root folder
						saveFile = new File(materialFolder + "/" + fileName);
					}
				} else {
					// This material has ≤10 parts or uses CutContour, no subfolders needed
					// ALL files go directly in the material folder
					saveFile = new File(materialFolder + "/" + fileName);
				}
			}
		} else {
			// Use the stored parent save path
			saveFile = new File(parentSavePath + "/" + fileName);
		}
		
		// Check if file exists
		if (saveFile.exists) {
			var overwrite = false;
			
			if (overwriteAll === null) {
				var dialog = confirm("File '" + fileName + "' already exists.\nDo you want to overwrite it?\n\n" +
								   "Click 'Yes' to overwrite, 'No' to skip.");
				
				if (dialog) {
					overwrite = true;
					// Ask if should apply to all
					if (confirm("Apply this decision to all existing files?")) {
						overwriteAll = true;
					}
				} else {
					if (confirm("Skip all existing files?")) {
						skipAll = true;
					}
				}
			} else {
				overwrite = overwriteAll;
			}
			
			if (!overwrite || skipAll) {
				return; // Skip this file
			}
		}
		
		// Set up PDF save options with Large Canvas compatibility
		var pdfOptions = new PDFSaveOptions();
		pdfOptions.compatibility = PDFCompatibility.ACROBAT7; // PDF 1.6 for Large Canvas support
		pdfOptions.generateThumbnails = true;
		pdfOptions.preserveEditability = true;
		pdfOptions.preset = "[Illustrator Default]";
		pdfOptions.viewAfterSaving = false;
		
		// Save as PDF directly
		tempDoc.saveAs(saveFile, pdfOptions);
		
		// Force completion
		app.redraw();
	}
    
    // Extract artboard name from filename for folder organization
    function getArtboardNameFromFileName(fileName) {
        // Format: basename_TYPE_artboardname.pdf
        var parts = fileName.split('_');
        if (parts.length >= 3) {
            var artboardName = parts.slice(2).join('_').replace('.pdf', '');
            return artboardName;
        }
        return "Unknown";
    }
    
    // Extract material name from artboard name for folder organization
    function getMaterialNameFromArtboardName(artboardName) {
        // First strip front/back designation
        var cleanName = stripFrontBackDesignation(artboardName);
        
        // Extract material name (everything before _Pt if present)
        var ptIndex = cleanName.toLowerCase().indexOf('_pt');
        if (ptIndex !== -1) {
            return cleanName.substring(0, ptIndex);
        } else {
            return cleanName;
        }
    }
    
    // Create material folder and return path
    function createMaterialFolder(materialName) {
        try {
            // Use the original case from materialData if available
            var folderName = materialName;
            if (materialData[materialName] && materialData[materialName].originalCase) {
                folderName = materialData[materialName].originalCase;
            }
            
            var materialFolder = new Folder(parentSavePath + "/" + folderName);
            if (!materialFolder.exists) {
                materialFolder.create();
            }
            return materialFolder.fsName;
        } catch (e) {
            return null; // Folder creation failed
        }
    }
    
    // Create subfolder (PRINT or CUT) and return path
    function createSubFolder(parentFolderPath, subFolderName) {
        try {
            var subFolder = new Folder(parentFolderPath + "/" + subFolderName);
            if (!subFolder.exists) {
                subFolder.create();
            }
            return subFolder.fsName;
        } catch (e) {
            return null; // Subfolder creation failed
        }
    }
    
// Generate filename with type prefix
    function generateFileName(type, artboardName) {
        // Use the stored original document path to get the name
        var originalName = originalDocPath.name;
        // Remove any file extension (.ai, .pdf, .eps, .svg)
        var baseName = originalName.replace(/\.(ai|pdf|eps|svg)$/i, "");
        baseName = baseName.replace("_PRIME", "");
        
        // Detect front/back designation for this artboard
        var frontBackDesignation = detectFrontBackDesignation(artboardName);
        var cleanArtboardName = stripFrontBackDesignation(artboardName);
        
        // Build filename with front/back designation if present
        var fileName = baseName + "_" + type + "_" + cleanArtboardName;
        if (frontBackDesignation) {
            fileName += "_" + frontBackDesignation;
        }
        fileName += ".pdf";
        
        return fileName;
    }
    
})();