/*
@METADATA
{
  "name": "Wrap Template Converter",
  "description": "Prepare Wrap Design Files From Templates",
  "version": "1.1",
  "target": "illustrator",
  "tags": ["wrap", "template", "processors"]
}
@END_METADATA
*/

#target illustrator

function processWrapTemplate() {
    // Check if a document is open
    if (app.documents.length === 0) {
        alert("Please open a document before running this script.");
        return;
    }
    
    var doc = app.activeDocument;
    
    try {
        // 1. Delete "Your Logo Here Image" - it's always the top layer
        var topLayer = doc.layers[0];
        if (topLayer.name.indexOf("Your Logo Here") !== -1 || topLayer.name.indexOf("Logo") !== -1) {
            topLayer.remove();
        }
        
        // 2. Find and process "Design on Layers Below" layer
        var designLayer = null;
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name.indexOf("Design on Layers Below") !== -1) {
                designLayer = doc.layers[i];
                break;
            }
        }
        
        if (designLayer) {
            // Clear all contents of this layer (including "Design Layer Image")
            clearLayerContents(designLayer);
            
            // Rename the layer to "Design"
            designLayer.name = "Design";
        }
        
        // 3. Rename artboard to "Proof"
        if (doc.artboards.length > 0) {
            doc.artboards[0].name = "Proof";
        }
        
        // 4. Lock "Template Layers | Do not edit!"
        var templateLayer = findLayerByName(doc.layers, "Template Layers | Do not edit!");
        if (templateLayer) {
            templateLayer.locked = true;
        }
        
        // 5. Convert document and all artwork to CMYK color mode
        if (doc.documentColorSpace !== DocumentColorSpace.CMYK) {
            try {
                // Use the convertToColorSpace method - this converts both document AND artwork
                doc.convertToColorSpace(DocumentColorSpace.CMYK);
                // Refresh document reference after color conversion
                doc = app.activeDocument;
            } catch (e) {
                // If that fails, try using the menu command
                try {
                    app.executeMenuCommand("doc-color-cmyk");
                    // Refresh document reference after color conversion
                    doc = app.activeDocument;
                } catch (e2) {
                    alert("Warning: Could not convert document to CMYK. Document will be saved in current color mode.");
                }
            }
        }
        
        // 6. Set active layer to "Design"
        var newDesignLayer = findLayerByName(doc.layers, "Design");
        if (newDesignLayer) {
            doc.activeLayer = newDesignLayer;
        }
        
        // 7. Save as PDF with layers preserved
        saveAsPDFWithLayers(doc);
        
        // 8. Close the file after saving
        doc.close(SaveOptions.DONOTSAVECHANGES);
        
        alert("Template processed successfully, saved as PDF, and file closed!");
        
    } catch (error) {
        alert("Error processing template: " + error.message);
    }
}

// Clear all contents of a layer (sublayers and page items)
function clearLayerContents(layer) {
    try {
        // Remove all sublayers
        if (layer.layers && layer.layers.length > 0) {
            for (var i = layer.layers.length - 1; i >= 0; i--) {
                layer.layers[i].remove();
            }
        }
        
        // Remove all page items (shapes, images, text, etc.)
        if (layer.pageItems && layer.pageItems.length > 0) {
            for (var j = layer.pageItems.length - 1; j >= 0; j--) {
                layer.pageItems[j].remove();
            }
        }
        
        return true;
    } catch (e) {
        return false;
    }
}

// Find layer by name (searches recursively)
function findLayerByName(layersCollection, targetName) {
    for (var i = 0; i < layersCollection.length; i++) {
        var layer = layersCollection[i];
        
        if (layer.name === targetName) {
            return layer;
        }
        
        if (layer.layers && layer.layers.length > 0) {
            var found = findLayerByName(layer.layers, targetName);
            if (found) {
                return found;
            }
        }
    }
    return null;
}

// Collapse all sublayers of a given layer
function collapseAllSublayers(parentLayer) {
    if (parentLayer.layers && parentLayer.layers.length > 0) {
        for (var i = 0; i < parentLayer.layers.length; i++) {
            var sublayer = parentLayer.layers[i];
            try {
                // Try to set preview to false (this sometimes works for collapsing)
                sublayer.preview = false;
            } catch (e) {
                // If that doesn't work, try using menu command
                try {
                    doc.activeLayer = sublayer;
                    // Unfortunately, there's no reliable way to collapse layers via script
                    // This is a limitation of Illustrator's scripting API
                } catch (e2) {
                    // Continue with next layer
                }
            }
            
            // Recursively collapse sublayers
            if (sublayer.layers && sublayer.layers.length > 0) {
                collapseAllSublayers(sublayer);
            }
        }
    }
}

// Save document as PDF using Illustrator Default preset
function saveAsPDFWithLayers(doc) {
    try {
        // Get the document's current file path and create PDF filename
        var docPath = doc.path;
        var docName = doc.name;
        
        // Remove file extension and add .pdf
        var baseName = docName.replace(/\.[^\/\.]+$/, "");
        var pdfFile = new File(docPath + "/" + baseName + ".pdf");
        
        // Use the Illustrator Default preset which preserves layers
        var pdfSaveOptions = new PDFSaveOptions();
        pdfSaveOptions.pdfPreset = "[Illustrator Default]";
        
        // Save the PDF
        doc.saveAs(pdfFile, pdfSaveOptions);
        
    } catch (error) {
        alert("Error saving PDF: " + error.message + "\nThe template has been processed but not saved as PDF.");
    }
}

// Run the script
processWrapTemplate();