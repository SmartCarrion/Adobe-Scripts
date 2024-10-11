/*
Instructions:
1. Select rectangular frames in your InDesign document.
2. Run this script (File > Scripts > Other Script...).
3. Select a folder containing your images (.jpg, .jpeg, .png, .tif, .tiff).
4. The script will process images in Photoshop and place them in InDesign.

Note: Photoshop must be installed on the same machine.
For OneDrive folders, ensure files are synced and available offline.
*/

#target indesign

function processAndPlaceImages() {
    if (app.documents.length === 0 || app.selection.length === 0) {
        alert("Please open a document and select some rectangular frames.");
        return;
    }

    // Photoshop processing function
    var psProcessingCode = function () {
        var inputFolder = Folder.selectDialog("Select the folder containing your images");
        if (inputFolder == null) return;

        var outputFolder = Folder.selectDialog("Select the folder to save processed images");
        if (outputFolder == null) return;

        function scanFolder(folder) {
            var imageFiles = [];
            var files = folder.getFiles();
            for (var i = 0; i < files.length; i++) {
                var file = files[i];
                if (file instanceof File && file.name.match(/\.(jpg|jpeg|png|tif|tiff)$/i)) {
                    imageFiles.push(file);
                } else if (file instanceof Folder) {
                    imageFiles = imageFiles.concat(scanFolder(file));
                }
            }
            return imageFiles;
        }

        var imageFiles = scanFolder(inputFolder);
        var processedFiles = [];

        for (var i = 0; i < imageFiles.length; i++) {
            try {
                app.open(imageFiles[i]);
                var doc = app.activeDocument;

                doc.changeMode(ChangeMode.GRAYSCALE);

                var levelLayer = doc.adjustmentLayers.add();
                levelLayer.kind = AdjustmentLayerType.LEVELS;
                levelLayer.adjustments.levels[0].input = [5, 1.0, 250];

                doc.flatten();

                var saveFile = new File(outputFolder + "/" + doc.name.replace(/\.[^\.]+$/, '') + "_BW.tif");
                var tiffSaveOptions = new TiffSaveOptions();
                tiffSaveOptions.imageCompression = TIFFEncoding.NONE;
                doc.saveAs(saveFile, tiffSaveOptions, true, Extension.LOWERCASE);

                processedFiles.push(saveFile.fsName);

                doc.close(SaveOptions.DONOTSAVECHANGES);
            } catch (e) {
                alert("Error processing file: " + imageFiles[i].name + "\nError: " + e);
            }
        }

        return processedFiles;
    };

    // Run Photoshop processing
    var bt = new BridgeTalk();
    bt.target = "photoshop";
    bt.body = psProcessingCode.toString() + "; psProcessingCode();";
    bt.onResult = function (resObj) {
        var processedFiles = eval(resObj.body);
        placeImagesInFrames(processedFiles);
    };
    bt.send();
}

function placeImagesInFrames(imageFiles) {
    var doc = app.activeDocument;
    var sel = app.selection;

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

    var allRectangles = [];
    for (var i = 0; i < sel.length; i++) {
        allRectangles = allRectangles.concat(findRectangles(sel[i]));
    }

    var availableImages = [].concat(imageFiles);
    var frameIndex = 0;

    for (var i = 0; i < allRectangles.length; i++) {
        var currentFrame = allRectangles[i];

        if (availableImages.length > 0) {
            var randomIndex = Math.floor(Math.random() * availableImages.length);
            var imageFile = availableImages[randomIndex];

            try {
                currentFrame.place(File(imageFile));
                currentFrame.fit(FitOptions.PROPORTIONALLY);
                currentFrame.fit(FitOptions.CENTER_CONTENT);
                frameIndex++;

                availableImages.splice(randomIndex, 1);
            } catch (e) {
                alert("Error placing image: " + e + "\nFile: " + imageFile + "\nContinuing to next image.");
                availableImages.splice(randomIndex, 1);
                i--;
            }
        } else if (imageFiles.length > 0) {
            availableImages = [].concat(imageFiles);
            i--;
        } else {
            alert("Ran out of images to place. Placed " + frameIndex + " images before stopping.");
            return;
        }
    }

    alert("Successfully placed " + frameIndex + " images.");
}

processAndPlaceImages();