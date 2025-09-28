/*
@METADATA
{
  "name": "Address Format Corrector",
  "description": "Corrects address formatting to proper postal standards",
  "version": "1.0",
  "target": "illustrator",
  "tags": ["address", "formatting", "postal", "text"]
}
@END_METADATA
*/

#target illustrator

(function() {
    // Enhanced cleanup - main context
    try {
        if (typeof addressPalette !== 'undefined' && addressPalette instanceof Window) {
            addressPalette.close();
        }
    } catch (e) {
        // Ignore cleanup errors
    }
    
    // Create palette window
    var addressPalette = new Window("palette", "Address Format Corrector");
    addressPalette.orientation = "column";
    addressPalette.alignChildren = "left";
    addressPalette.preferredSize.width = 400;
    addressPalette.spacing = 10;
    addressPalette.margins = 15;
    
    // Title
    var title = addressPalette.add("statictext", undefined, "Address Format Corrector");
    title.graphics.font = ScriptUI.newFont("Arial", "BOLD", 12);
    
    // Input section
    var inputGroup = addressPalette.add("group");
    inputGroup.orientation = "column";
    inputGroup.alignChildren = "left";
    inputGroup.spacing = 5;
    
    inputGroup.add("statictext", undefined, "Original address (or select text in document):");
    var inputField = inputGroup.add("edittext", undefined, "", {multiline: true});
    inputField.preferredSize.width = 370;
    inputField.preferredSize.height = 60;
    
    // Output section
    var outputGroup = addressPalette.add("group");
    outputGroup.orientation = "column";
    outputGroup.alignChildren = "left";
    outputGroup.spacing = 5;
    
    outputGroup.add("statictext", undefined, "Corrected address:");
    var outputField = outputGroup.add("edittext", undefined, "", {multiline: true});
    outputField.preferredSize.width = 370;
    outputField.preferredSize.height = 60;
    
    // Options section
    var optionsGroup = addressPalette.add("group");
    optionsGroup.orientation = "column";
    optionsGroup.alignChildren = "left";
    optionsGroup.spacing = 5;
    
    optionsGroup.add("statictext", undefined, "Options:");
    var uppercaseCheck = optionsGroup.add("checkbox", undefined, "Convert city/state/zip to UPPERCASE");
    uppercaseCheck.value = true;
    
    // Buttons
    var buttonGroup = addressPalette.add("group");
    buttonGroup.alignment = "center";
    buttonGroup.spacing = 10;
    
    var formatButton = buttonGroup.add("button", undefined, "Format Address");
    var resetButton = buttonGroup.add("button", undefined, "Reset");
    var closeButton = buttonGroup.add("button", undefined, "Close");
    
    // State abbreviations lookup
    var stateAbbreviations = {
        "alabama": "AL", "alaska": "AK", "arizona": "AZ", "arkansas": "AR", "california": "CA",
        "colorado": "CO", "connecticut": "CT", "delaware": "DE", "florida": "FL", "georgia": "GA",
        "hawaii": "HI", "idaho": "ID", "illinois": "IL", "indiana": "IN", "iowa": "IA",
        "kansas": "KS", "kentucky": "KY", "louisiana": "LA", "maine": "ME", "maryland": "MD",
        "massachusetts": "MA", "michigan": "MI", "minnesota": "MN", "mississippi": "MS", "missouri": "MO",
        "montana": "MT", "nebraska": "NE", "nevada": "NV", "new hampshire": "NH", "new jersey": "NJ",
        "new mexico": "NM", "new york": "NY", "north carolina": "NC", "north dakota": "ND", "ohio": "OH",
        "oklahoma": "OK", "oregon": "OR", "pennsylvania": "PA", "rhode island": "RI", "south carolina": "SC",
        "south dakota": "SD", "tennessee": "TN", "texas": "TX", "utah": "UT", "vermont": "VT",
        "virginia": "VA", "washington": "WA", "west virginia": "WV", "wisconsin": "WI", "wyoming": "WY",
        "district of columbia": "DC", "puerto rico": "PR"
    };
    
    // Street type abbreviations
    var streetTypes = {
        "street": "St", "avenue": "Ave", "boulevard": "Blvd", "drive": "Dr", "road": "Rd",
        "lane": "Ln", "court": "Ct", "circle": "Cir", "place": "Pl", "way": "Way",
        "parkway": "Pkwy", "highway": "Hwy", "trail": "Trl", "terrace": "Ter", "square": "Sq"
    };
    
    // Direction abbreviations
    var directions = {
        "north": "N", "south": "S", "east": "E", "west": "W",
        "northeast": "NE", "northwest": "NW", "southeast": "SE", "southwest": "SW"
    };
    
    // Format address function
    function formatAddress(input, useUppercase) {
        if (!input || input === "") {
            return { success: false, error: "Please enter an address" };
        }
        
        var lines = input.split(/[\n\r]+/);
        var cleanLines = [];
        
        // Clean and collect non-empty lines
        for (var i = 0; i < lines.length; i++) {
            var line = lines[i].replace(/^\s+|\s+$/g, ''); // trim
            if (line !== "") {
                cleanLines.push(line);
            }
        }
        
        if (cleanLines.length === 0) {
            return { success: false, error: "No valid address content found" };
        }
        
        // Check if this is a single line that should be formatted as USPS single-line
        if (cleanLines.length === 1) {
            return formatSingleLineAddress(cleanLines[0], useUppercase);
        }
        
        // Multi-line address - format each line appropriately
        var formattedLines = [];
        
        for (var i = 0; i < cleanLines.length; i++) {
            var line = cleanLines[i];
            var formattedLine = line;
            
            // Check if this looks like a city/state/zip line (last line typically)
            var cityStateZipPattern = /^(.+?),?\s*([A-Za-z\s]+)\s+(\d{5}(?:-\d{4})?)(?:\s*(.*))?$/;
            var match = line.match(cityStateZipPattern);
            
            if (match && i === cleanLines.length - 1) {
                // This is likely the city/state/zip line
                var city = match[1].replace(/,\s*$/, '').replace(/^\s+|\s+$/g, '');
                var state = match[2].replace(/^\s+|\s+$/g, '');
                var zip = match[3];
                var extra = match[4] ? match[4].replace(/^\s+|\s+$/g, '') : '';
                
                // Convert state to abbreviation if it's a full name
                var stateLower = state.toLowerCase();
                if (stateAbbreviations[stateLower]) {
                    state = stateAbbreviations[stateLower];
                }
                
                if (useUppercase) {
                    city = city.toUpperCase();
                    state = state.toUpperCase();
                } else {
                    // Title case for city
                    city = titleCase(city);
                    state = state.toUpperCase(); // States are always uppercase
                }
                
                formattedLine = city + ", " + state + " " + zip;
                if (extra) {
                    formattedLine += " " + extra;
                }
            } else {
                // This is likely a street address line
                formattedLine = formatStreetAddress(line);
            }
            
            formattedLines.push(formattedLine);
        }
        
        return {
            success: true,
            formatted: formattedLines.join("\n")
        };
    }
    
    // Format single-line address according to USPS standard
    function formatSingleLineAddress(line, useUppercase) {
        // Try to parse the single line into components
        // Look for patterns like: Name, Address, City State Zip
        // or: Address City State Zip
        
        var parts = line.split(/,\s*/);
        var result = "";
        
        if (parts.length >= 3) {
            // Likely format: Name, Address, City State Zip
            var name = parts[0].replace(/^\s+|\s+$/g, '');
            var address = parts[1].replace(/^\s+|\s+$/g, '');
            var cityStateZip = parts[2].replace(/^\s+|\s+$/g, '');
            
            // Parse city, state, zip from the last part
            var cityStateZipMatch = cityStateZip.match(/^(.+?)\s+([A-Za-z]{2})\s+(\d{5}(?:-\d{4})?)$/);
            if (!cityStateZipMatch) {
                // Try with full state name
                cityStateZipMatch = cityStateZip.match(/^(.+?)\s+([A-Za-z\s]+)\s+(\d{5}(?:-\d{4})?)$/);
            }
            
            if (cityStateZipMatch) {
                var city = cityStateZipMatch[1].replace(/^\s+|\s+$/g, '');
                var state = cityStateZipMatch[2].replace(/^\s+|\s+$/g, '');
                var zip = cityStateZipMatch[3];
                
                // Convert state to abbreviation if it's a full name
                var stateLower = state.toLowerCase();
                if (stateAbbreviations[stateLower]) {
                    state = stateAbbreviations[stateLower];
                }
                
                // Format according to USPS single-line standard
                if (useUppercase) {
                    result = name.toUpperCase() + ", " + formatStreetAddress(address).toUpperCase() + ", " + city.toUpperCase() + ", " + state.toUpperCase() + " " + zip;
                } else {
                    result = titleCase(name) + ", " + formatStreetAddress(address) + ", " + titleCase(city) + ", " + state.toUpperCase() + " " + zip;
                }
            } else {
                // Fallback - just format the address part
                result = titleCase(name) + ", " + formatStreetAddress(address) + ", " + cityStateZip;
            }
        } else if (parts.length === 2) {
            // Likely format: Address, City State Zip
            var address = parts[0].replace(/^\s+|\s+$/g, '');
            var cityStateZip = parts[1].replace(/^\s+|\s+$/g, '');
            
            // Parse city, state, zip
            var cityStateZipMatch = cityStateZip.match(/^(.+?)\s+([A-Za-z]{2})\s+(\d{5}(?:-\d{4})?)$/);
            if (!cityStateZipMatch) {
                cityStateZipMatch = cityStateZip.match(/^(.+?)\s+([A-Za-z\s]+)\s+(\d{5}(?:-\d{4})?)$/);
            }
            
            if (cityStateZipMatch) {
                var city = cityStateZipMatch[1].replace(/^\s+|\s+$/g, '');
                var state = cityStateZipMatch[2].replace(/^\s+|\s+$/g, '');
                var zip = cityStateZipMatch[3];
                
                var stateLower = state.toLowerCase();
                if (stateAbbreviations[stateLower]) {
                    state = stateAbbreviations[stateLower];
                }
                
                if (useUppercase) {
                    result = formatStreetAddress(address).toUpperCase() + ", " + city.toUpperCase() + ", " + state.toUpperCase() + " " + zip;
                } else {
                    result = formatStreetAddress(address) + ", " + titleCase(city) + ", " + state.toUpperCase() + " " + zip;
                }
            } else {
                result = formatStreetAddress(address) + ", " + cityStateZip;
            }
        } else {
            // Single part - try to parse as address + city state zip
            var addressMatch = line.match(/^(.+?)\s+([A-Za-z\s]+)\s+([A-Za-z]{2})\s+(\d{5}(?:-\d{4})?)$/);
            if (!addressMatch) {
                addressMatch = line.match(/^(.+?)\s+([A-Za-z\s]+)\s+([A-Za-z\s]+)\s+(\d{5}(?:-\d{4})?)$/);
            }
            
            if (addressMatch) {
                var address = addressMatch[1];
                var city = addressMatch[2];
                var state = addressMatch[3];
                var zip = addressMatch[4];
                
                var stateLower = state.toLowerCase();
                if (stateAbbreviations[stateLower]) {
                    state = stateAbbreviations[stateLower];
                }
                
                if (useUppercase) {
                    result = formatStreetAddress(address).toUpperCase() + ", " + city.toUpperCase() + ", " + state.toUpperCase() + " " + zip;
                } else {
                    result = formatStreetAddress(address) + ", " + titleCase(city) + ", " + state.toUpperCase() + " " + zip;
                }
            } else {
                // Can't parse - just clean up what we have
                result = formatStreetAddress(line);
            }
        }
        
        return {
            success: true,
            formatted: result
        };
    }
    
    // Format street address
    function formatStreetAddress(line) {
        var words = line.split(/\s+/);
        var formattedWords = [];
        
        for (var i = 0; i < words.length; i++) {
            var word = words[i].toLowerCase();
            var originalWord = words[i];
            
            // Check for directions
            if (directions[word]) {
                formattedWords.push(directions[word]);
            }
            // Check for street types
            else if (streetTypes[word]) {
                formattedWords.push(streetTypes[word]);
            }
            // Check if it's a number (keep original formatting)
            else if (/^\d/.test(originalWord)) {
                formattedWords.push(originalWord);
            }
            // Regular words - title case
            else {
                formattedWords.push(titleCase(originalWord));
            }
        }
        
        return formattedWords.join(" ");
    }
    
    // Title case function
    function titleCase(str) {
        return str.toLowerCase().replace(/\b\w/g, function(l) {
            return l.toUpperCase();
        });
    }
    
    // Format button handler - using BridgeTalk for document access
    formatButton.onClick = function() {
        // 1. Read the input from the text field FIRST.
        var manualInput = inputField.text;
        
        // 2. Clear the output field immediately (synchronously).
        outputField.text = "";
        
        var bt_read = new BridgeTalk();
        bt_read.target = "illustrator";
        var scriptToRunInIllustrator = "var selectedText = null; try { if (app.documents.length > 0 && app.activeDocument.selection.length > 0) { var sel = app.activeDocument.selection; for (var i = 0; i < sel.length; i++) { if (sel[i].typename === 'TextFrame') { selectedText = sel[i].contents; break; } } } } catch(e) {}; selectedText;";
        bt_read.body = scriptToRunInIllustrator;
        
        bt_read.onResult = function(resultObj) {
            var selectedText = resultObj.body;
            var input = "";
            var usedSelection = false;
            
            if (selectedText && selectedText !== "null") {
                input = selectedText;
                inputField.text = input;
                usedSelection = true;
            } else {
                // Fallback to the manually entered text we saved.
                input = manualInput;
            }
            
            if (!input || input.replace(/^\s+|\s+$/g, '') === "") {
                outputField.text = "Error: Please enter an address";
                return;
            }
            
            var result = formatAddress(input, uppercaseCheck.value);
            
            if (result.success) {
                // 3. THE FIX: Assign ONLY to the .text property.
                outputField.text = result.formatted;
                outputField.active = true; // We can still focus the field.
                
                // If input came from selected text, update the text object
                if (usedSelection) {
                    var newContent = result.formatted;
                    // Escape single quotes and backslashes for the script
                    var escapedContent = newContent.replace(/\\/g, '\\\\').replace(/'/g, "\\'").replace(/\r\n/g, '\\r').replace(/\n/g, '\\r').replace(/\r/g, '\\r');
                    
                    var scriptToWrite = "try { if (app.documents.length > 0 && app.activeDocument.selection.length > 0 && app.activeDocument.selection[0].typename === 'TextFrame') { app.activeDocument.selection[0].contents = '" + escapedContent + "'; } } catch(e) {}";
                    
                    var bt_write = new BridgeTalk();
                    bt_write.target = "illustrator";
                    bt_write.body = scriptToWrite;
                    bt_write.send();
                }
            } else {
                outputField.text = "Error: " + result.error;
            }
        }
        
        bt_read.send();
    };
    
    // Reset button handler
    resetButton.onClick = function() {
        inputField.text = "";
        outputField.text = "";
        inputField.active = true;
    };
    
    // Close button handler
    closeButton.onClick = function() {
        addressPalette.close();
    };
    
    // Enter key triggers formatting
    addressPalette.addEventListener("keydown", function(event) {
        if (event.keyName === "Enter") {
            formatButton.notify();
        }
    });
    
    // Clear output when input changes
    inputField.onChanging = function() {
        outputField.text = "";
    };
    
    // Show the palette
    addressPalette.show();
    this.addressPalette = addressPalette;
    
})();