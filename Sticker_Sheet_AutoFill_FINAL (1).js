#target photoshop
app.bringToFront();

// =====================================================
// STICKER SHEET AUTO FILL SCRIPT
// - Top margin: 5mm
// - Gap between stickers: 15mm
// - Direction: TOP → BOTTOM
// - Auto canvas height
// - Original sticker size (no resize)
// =====================================================

// =======================
// SETTINGS
// =======================
var TOP_MARGIN_MM = 5;   // gap at the start of the sheet
var GAP_MM = 15;        // spacing between stickers
var MM_TO_PX = 2.83464567; // mm to px conversion (Photoshop scripting base)

var TOP_MARGIN_PX = TOP_MARGIN_MM * MM_TO_PX;
var GAP_PX = GAP_MM * MM_TO_PX;

// =======================
// CHECK DOCUMENT
// =======================
if (app.documents.length === 0) {
    alert("Please open a blank document before running the script.");
    exit();
}

var doc = app.activeDocument;
var startY = TOP_MARGIN_PX;

// =======================
// ASK STICKER TYPE COUNT
// =======================
var stickerTypeCount = prompt("How many sticker types do you want to add?", "1");
stickerTypeCount = parseInt(stickerTypeCount, 10);

if (isNaN(stickerTypeCount) || stickerTypeCount < 1) {
    alert("Invalid number of sticker types.");
    exit();
}

// =======================
// MAIN PROCESS
// =======================
for (var i = 0; i < stickerTypeCount; i++) {

    alert("Select TIF file for Sticker " + (i + 1));
    var file = File.openDialog("Select sticker TIF file", "*.tif");

    if (!file) {
        alert("File selection cancelled.");
        exit();
    }

    var repeatCount = prompt("How many times do you want to repeat this sticker?", "1");
    repeatCount = parseInt(repeatCount, 10);

    if (isNaN(repeatCount) || repeatCount < 1) {
        alert("Invalid repeat count.");
        exit();
    }

    for (var r = 0; r < repeatCount; r++) {

        // Open sticker file
        var stickerDoc = app.open(file);
        var stickerLayer = stickerDoc.activeLayer;

        // Duplicate sticker into main document
        stickerLayer.duplicate(doc, ElementPlacement.PLACEATBEGINNING);
        stickerDoc.close(SaveOptions.DONOTSAVECHANGES);

        var placedLayer = doc.activeLayer;

        // Sticker dimensions
        var bounds = placedLayer.bounds;
        var stickerWidth  = bounds[2].as("px") - bounds[0].as("px");
        var stickerHeight = bounds[3].as("px") - bounds[1].as("px");

        // Extend canvas height automatically if required
        var requiredHeight = startY + stickerHeight;
        if (requiredHeight > doc.height.as("px")) {
            doc.resizeCanvas(
                doc.width,
                requiredHeight + GAP_PX,
                AnchorPosition.TOPCENTER
            );
        }

        // Calculate target position (centered horizontally)
        var targetX = (doc.width.as("px") - stickerWidth) / 2;
        var targetY = startY;

        placedLayer.translate(
            targetX - bounds[0].as("px"),
            targetY - bounds[1].as("px")
        );

        // Move cursor down for next sticker
        startY += stickerHeight + GAP_PX;
    }
}

alert("Sticker sheet created successfully!");
