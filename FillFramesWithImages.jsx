/*
Instructions:
1. Select rectangular frames or groups containing rectangles in your InDesign document.
2. Run this script (File > Scripts > Other Script...).
3. Select a folder containing your images (.tif, .jpg, .jpeg, .png).
4. The script will randomly place images in the selected frames.

located in:  C:\Users\<username>\AppData\Roaming\Adobe\InDesign\Version 19.0\en_US\Scripts\Scripts Panel
generally it would be put in the Adobe Scripts folder:
//Windows: C:\Program Files\Adobe\Adobe InDesign [version]\Presets\Scripts\
//and macOS: /Applications/Adobe InDesign [version]/Presets/Scripts/

Note: This script works best with local folders. For OneDrive, ensure files are available offline.
*/

function placeImagesInFrames() {
    // Show instructions
    var instructions = "Instructions:\n\n" +
        "1. Select rectangular frames or groups containing rectangles.\n" +
        "2. Run this script.\n" +
        "3. Select a folder containing your images (.tif, .jpg, .jpeg, .png).\n" +
        "4. Images will be randomly placed in the selected frames.\n\n" +
        "Note: For OneDrive folders, ensure files are available offline.";
    alert(instructions);

    if (app.documents.length > 0 && app.selection.length > 0) {
        var doc = app.activeDocument;
        var sel = app.selection;

        // Prompt user to select the image folder
        var imageFolder = Folder.selectDialog("Select the folder containing your images");

        if (imageFolder != null) {
            // Check if it's a OneDrive path
            if (imageFolder.fsName.indexOf("OneDrive") !== -1) {
                var oneDriveWarning = "You've selected a OneDrive folder. If you encounter errors, please ensure all files are available offline or copy them to a local folder.";
                alert(oneDriveWarning);
            }

            // Get all image files (TIFF, JPEG, PNG)
            var imageFiles = imageFolder.getFiles(function (file) {
                return file.name.match(/\.(tif|tiff|jpg|jpeg|png)$/i) != null;
            });
            var availableImages = [].concat(imageFiles);  // Create a copy of the array
            var frameIndex = 0;

            // Function to recursively find rectangles in groups
            function findRectangles(item) {
                var rectangles = [];
                if (item instanceof Rectangle) {
                    rectangles.push(item);
                } else if (item instanceof Group) {
                    for (var i = 0; i < item.allPageItems.length; i++) {
                        rectangles = rectangles.concat(findRectangles(item.allPageItems[i]));
                    }
                }
                return rectangles;
            }

            // Collect all rectangles from selection, including those in groups
            var allRectangles = [];
            for (var i = 0; i < sel.length; i++) {
                allRectangles = allRectangles.concat(findRectangles(sel[i]));
            }

            // Place images in found rectangles
            for (var i = 0; i < allRectangles.length; i++) {
                var currentFrame = allRectangles[i];

                if (availableImages.length > 0) {
                    // Randomly select an image
                    var randomIndex = Math.floor(Math.random() * availableImages.length);
                    var imageFile = availableImages[randomIndex];

                    try {
                        currentFrame.place(imageFile);
                        currentFrame.fit(FitOptions.PROPORTIONALLY);  // Fit proportionally
                        currentFrame.fit(FitOptions.CENTER_CONTENT);  // Center the content
                        frameIndex++;

                        // Remove the used image from the available list
                        availableImages.splice(randomIndex, 1);
                    } catch (e) {
                        alert("Error placing image: " + e + "\nFile: " + imageFile.name + "\nStopping script execution.");
                        return; // Stop the script on error
                    }
                } else if (imageFiles.length > 0) {
                    // If we've used all images but still have frames, start over
                    availableImages = [].concat(imageFiles);
                    i--;  // Retry this frame
                } else {
                    alert("Ran out of images to place. Placed " + frameIndex + " images before stopping.");
                    return; // Stop the script when out of images
                }
            }

            alert("Successfully placed " + frameIndex + " images.");
        }
    } else {
        alert("Please open a document and select some rectangular frames or groups containing rectangles.");
    }
}

// Execute the function
placeImagesInFrames();