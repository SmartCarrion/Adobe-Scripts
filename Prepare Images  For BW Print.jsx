// Photoshop script to prepare images for black and white printing

#target photoshop   
// this script is located in: C:\Users\imccl\AppData\Roaming\Adobe\Adobe Photoshop 2024\Presets\Scripts
//generally it would be put in the Adobe Scripts folder: 
//Windows: C:\Program Files\Adobe\Adobe Photoshop [version]\Presets\Scripts\
//and macOS: /Applications/Adobe Photoshop [version]/Presets/Scripts/

// User instructions
alert("This script will process images for black and white printing.\n\n" +
    "It will:\n" +
    "1. Convert images to grayscale\n" +
    "2. Adjust levels to ensure true black and white\n" +
    "3. Preserve gray tones and anti-aliased edges\n\n" +
    "Please select the folder containing your images.");

// Prompt user to select input folder
var inputFolder = Folder.selectDialog("Select the folder containing your images");

if (inputFolder != null) {
    // Prompt user to select output folder
    var outputFolder = Folder.selectDialog("Select the folder to save processed images");

    if (outputFolder != null) {
        // Get all image files in the input folder
        var fileList = inputFolder.getFiles(/\.(jpg|jpeg|png|tif|tiff)$/i);

        for (var i = 0; i < fileList.length; i++) {
            app.open(fileList[i]);
            var doc = app.activeDocument;

            // Convert to grayscale
            doc.changeMode(ChangeMode.GRAYSCALE);

            // Adjust levels
            var levelLayer = doc.adjustmentLayers.add();
            levelLayer.kind = AdjustmentLayerType.LEVELS;

            // Set black point to 5 and white point to 250
            levelLayer.adjustments.levels[0].input = [5, 1.0, 250];

            // Flatten the image
            doc.flatten();

            // Save as TIFF
            var saveFile = new File(outputFolder + "/" + doc.name.replace(/\.[^\.]+$/, '') + "_BW.tif");
            var tiffSaveOptions = new TiffSaveOptions();
            tiffSaveOptions.imageCompression = TIFFEncoding.NONE;
            doc.saveAs(saveFile, tiffSaveOptions, true, Extension.LOWERCASE);

            // Close the document
            doc.close(SaveOptions.DONOTSAVECHANGES);
        }

        alert("Processing complete. Processed " + fileList.length + " images.");
    }
}