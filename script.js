function exportSingleSlides(filepath){
	var app = new ActiveXObject("PowerPoint.Application");
	app.Visible = true;

	// Open the presentation
	var presentation = app.Presentations.Open(filepath, true);
	var tmpPresentation; // Used to store a new presentation for each slide.

	var e = new Enumerator(presentation.Slides);
	var slide;
	e.moveFirst();
	var i = 0;
	while (!e.atEnd()) {
		i++;
		slide = e.item(); // gets this slide

		// Variable for the unix timestamp. Used in file names below
		var unix = Math.round(+new Date()/1000);

		// Export slide to png
		slide.Export(scriptpath + 'TEMP\\' + 'slide' + unix + '.png', 'PNG', 1920, 1080);

		// Open new presentation
		tmpPresentation = app.Presentations.Add();

		// Set up the slide size to be the same as the source.
		tmpPresentation.PageSetup.SlideHeight = presentation.PageSetup.SlideHeight;
		tmpPresentation.PageSetup.SlideWidth = presentation.PageSetup.SlideWidth;

		// Get the layout from the source slide
		layout = slide.CustomLayout;
		
		// Copy current slide and paste into new presentation
		slide.Copy();
		tmpPresentation.Slides.Paste(1);

		// Set the layout. The line below won't work in the newest version of powerpoint for some reason.
		// tmpPresentation.Slides(1).CustomLayout = layout;
		
		// Save and close
		// SaveAs with two arguments: filename and filetype. See here for filetypes: http://www.bettersolutions.com/powerpoint/PRV283/LY621621612.htm
		// SaveAs() gives an "Unspecified error" when using a newer powerpoint version. I have no clue why.
		tmpPresentation.SaveAs(scriptpath + 'TEMP\\' + 'slide' + unix, 1);
		tmpPresentation.Close();
		e.moveNext();
	}

	// Close the presentation
	presentation.Close();
	app.Quit();
}

// The path to the script hardcoded (can't see why it would hurt here)
scriptpath = "C:\\Users\\Stian\\Dropbox\\jobb\\RB\\powerpoint\\";

// Array with all powerpoint files we need
filearray = ['C:\\Users\\Stian\\Dropbox\\jobb\\RB\\powerpoint\\test.pptx','C:\\Users\\Stian\\Dropbox\\jobb\\RB\\powerpoint\\test2.pptx'];

// Now we can loop over the powerpoint-files
for (i = 0; i < filearray.length; i++) {
	exportSingleSlides(filearray[i]);
}