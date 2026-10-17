// Exports the extension icons from visuals/icons_teams_excavator.ai.
// Usage: open the .ai in Vectorpea on vectorpea.com (the packaged app runs the
// script but never offers the files; Photopea works too), File > Script, paste
// this file, Run. It downloads output.zip with icon128/96/48/32.png: extract it
// into icons/. Script errors only show in the devtools console (F12).
(function () {
	// Descending order: each resize starts from the previous, larger state, so
	// any rasterized content is only ever scaled down.
	var SIZES = [128, 96, 48, 32];
	var doc = app.activeDocument;
	var ratio = +doc.height / +doc.width;

	var opts = new ExportOptionsSaveForWeb();
	opts.format = SaveDocumentType.PNG;
	opts.PNG8 = false;
	opts.transparency = true;

	for (var i = 0; i < SIZES.length; i++) {
		var size = SIZES[i];
		doc.resizeImage(size, Math.round(size * ratio));
		doc.exportDocument(new File("icon" + size + ".png"), ExportType.SAVEFORWEB, opts);
	}
})();
