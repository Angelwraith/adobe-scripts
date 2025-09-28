/*
@METADATA
{
  "name": "Auto-Group Overlapping",
  "description": "Create Groups From Isolated Overlapping Objects",
  "version": "1.0",
  "target": "illustrator",
  "tags": ["group", "auto", "utility"]
}
@END_METADATA
*/


(function() {
    'use strict';
    
    // Configuration
    var MAX_RECOMMENDED = 500; // Increased from 120 due to optimizations
    var BATCH_SIZE = 100; // Process in batches for very large selections
    
    // Validation
    if (!app.activeDocument) {
        alert('No active document found.');
        return;
    }
    
    if (app.selection.length < 2) {
        alert('Please select at least 2 objects to group.');
        return;
    }
    
    var selection = app.selection;
    var selectionLength = selection.length;
    
    // Warning for large selections
    if (selectionLength > MAX_RECOMMENDED) {
        var proceed = confirm(
            'You have selected ' + selectionLength + ' objects.\n' +
            'For optimal performance, consider selecting fewer than ' + MAX_RECOMMENDED + ' objects.\n' +
            'Do you want to continue?'
        );
        if (!proceed) return;
    }
    
    var startTime = new Date().getTime();
    
    try {
        var groups = processSelection(selection);
        var endTime = new Date().getTime();
        var processingTime = (endTime - startTime) / 1000;
        
        showResults(groups, selectionLength, processingTime);
        
    } catch (error) {
        alert('Error processing selection: ' + error.message);
    }
    
    /**
     * Main processing function
     */
    function processSelection(objects) {
        var objectCount = objects.length;
        var unionFind = new UnionFind(objectCount);
        var boundingBoxes = cacheBoundingBoxes(objects);
        
        // Find all overlapping pairs
        findOverlappingPairs(boundingBoxes, unionFind);
        
        // Get groups from Union-Find structure
        var groupMap = unionFind.getGroups();
        var groups = createGroups(objects, groupMap);
        
        return groups;
    }
    
    /**
     * Cache bounding boxes for performance
     */
    function cacheBoundingBoxes(objects) {
        var boxes = [];
        for (var i = 0; i < objects.length; i++) {
            var bounds = objects[i].geometricBounds;
            boxes[i] = {
                left: bounds[0],
                top: bounds[1],
                right: bounds[2],
                bottom: bounds[3],
                width: bounds[2] - bounds[0],
                height: bounds[1] - bounds[3],
                index: i
            };
        }
        return boxes;
    }
    
    /**
     * Find all overlapping pairs using optimized overlap detection
     */
    function findOverlappingPairs(boxes, unionFind) {
        var len = boxes.length;
        
        for (var i = 0; i < len - 1; i++) {
            for (var j = i + 1; j < len; j++) {
                if (boxesOverlap(boxes[i], boxes[j])) {
                    unionFind.union(i, j);
                }
            }
        }
    }
    
    /**
     * Optimized bounding box overlap detection
     */
    function boxesOverlap(box1, box2) {
        // Early termination - check if boxes are completely separate
        if (box1.right < box2.left || box2.right < box1.left) return false;
        if (box1.bottom > box2.top || box2.bottom > box1.top) return false;
        
        return true;
    }
    
    /**
     * Create actual Illustrator groups
     */
    function createGroups(objects, groupMap) {
        var groupCount = 0;
        var createdGroups = [];
        
        // Clear selection first
        app.selection = null;
        
        for (var rootId in groupMap) {
            var indices = groupMap[rootId];
            if (indices.length > 1) {
                var group = createSingleGroup(objects, indices);
                if (group) {
                    createdGroups.push(group);
                    groupCount++;
                }
            }
        }
        
        // Select all created groups
        if (createdGroups.length > 0) {
            app.selection = createdGroups;
        }
        
        app.redraw();
        return groupCount;
    }
    
    /**
     * Create a single group from objects at given indices
     */
    function createSingleGroup(objects, indices) {
        if (indices.length < 2) return null;
        
        try {
            // Find the frontmost object to determine group position
            var frontmostIndex = findFrontmostObject(objects, indices);
            var frontmostObject = objects[frontmostIndex];
            
            // Create new group
            var group = frontmostObject.parent.groupItems.add();
            group.move(frontmostObject, ElementPlacement.PLACEBEFORE);
            
            // Move objects to group (in reverse order to maintain stacking)
            for (var i = indices.length - 1; i >= 0; i--) {
                objects[indices[i]].move(group, ElementPlacement.INSIDE);
            }
            
            return group;
            
        } catch (error) {
            // If grouping fails, continue with other groups
            return null;
        }
    }
    
    /**
     * Find the frontmost object (lowest zOrderPosition)
     */
    function findFrontmostObject(objects, indices) {
        var frontmostIndex = indices[0];
        var frontmostZ = getZOrder(objects[frontmostIndex]);
        
        for (var i = 1; i < indices.length; i++) {
            var currentZ = getZOrder(objects[indices[i]]);
            if (currentZ < frontmostZ) {
                frontmostZ = currentZ;
                frontmostIndex = indices[i];
            }
        }
        
        return frontmostIndex;
    }
    
    /**
     * Safely get z-order position
     */
    function getZOrder(obj) {
        try {
            return obj.zOrderPosition || 0;
        } catch (e) {
            return 0;
        }
    }
    
    /**
     * Display results to user
     */
    function showResults(groupCount, totalObjects, processingTime) {
        var message = groupCount + ' groups created.\n' +
                     '------------------------\n' +
                     'Objects processed: ' + totalObjects + '\n' +
                     'Processing time: ' + processingTime.toFixed(2) + ' seconds';
        
        alert(message);
    }
    
    /**
     * Union-Find (Disjoint Set) data structure for efficient grouping
     */
    function UnionFind(size) {
        var self = this;
        self.parent = [];
        self.rank = [];
        
        // Initialize each element as its own parent
        for (var i = 0; i < size; i++) {
            self.parent[i] = i;
            self.rank[i] = 0;
        }
        
        self.find = function(x) {
            // Path compression optimization
            if (self.parent[x] !== x) {
                self.parent[x] = self.find(self.parent[x]);
            }
            return self.parent[x];
        };
        
        self.union = function(x, y) {
            var rootX = self.find(x);
            var rootY = self.find(y);
            
            if (rootX === rootY) return;
            
            // Union by rank optimization
            if (self.rank[rootX] < self.rank[rootY]) {
                self.parent[rootX] = rootY;
            } else if (self.rank[rootX] > self.rank[rootY]) {
                self.parent[rootY] = rootX;
            } else {
                self.parent[rootY] = rootX;
                self.rank[rootX]++;
            }
        };
        
        self.getGroups = function() {
            var groups = {};
            
            for (var i = 0; i < self.parent.length; i++) {
                var root = self.find(i);
                if (!groups[root]) {
                    groups[root] = [];
                }
                groups[root].push(i);
            }
            
            return groups;
        };
    }
    
})();