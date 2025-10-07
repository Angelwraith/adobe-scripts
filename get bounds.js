function getVisibleClippingMaskBounds(clippingGroup) {
  if (clippingGroup.typename !== "GroupItem" || !clippingGroup.clipped) {
    return null; // Not a clipping mask group
  }

  var doc = app.activeDocument;
  var originalAbIndex = doc.artboards.getActiveArtboardIndex();
  var originalAb = doc.artboards[originalAbIndex];
  var originalAbRect = originalAb.artboardRect;

  // 1. Create a temporary artboard to get the dimensions.
  var tempAb = doc.artboards.add(clippingGroup.geometricBounds);
  tempAb.name = "temp_bounds_calc";
  doc.artboards.setActiveArtboardIndex(doc.artboards.length - 1);

  // 2. Fit the temporary artboard to the artwork, which respects the mask.
  doc.artboards[doc.artboards.length - 1].fitToArt(clippingGroup.pageItems[0].visibleBounds);

  // 3. Get the bounds from the temporary artboard.
  var bounds = doc.artboards[doc.artboards.length - 1].artboardRect;
  var width = bounds[2] - bounds[0];
  var height = bounds[1] - bounds[3]; // Illustrator's Y-axis is inverted

  // 4. Clean up: restore original artboard and remove the temporary one.
  doc.artboards.setActiveArtboardIndex(originalAbIndex);
  doc.artboards[doc.artboards.length - 1].remove();

  return {
    width: width,
    height: height
  };
}


// Example usage:
if (app.selection.length > 0) {
  var selectedItem = app.selection[0];
  if (selectedItem.typename === "GroupItem" && selectedItem.clipped) {
    var dimensions = getVisibleClippingMaskBounds(selectedItem);
    if (dimensions) {
      alert("Visible Width: " + dimensions.width + "\n" + "Visible Height: " + dimensions.height);
    }
  } else {
    alert("Please select a clipping mask group.");
  }
}
