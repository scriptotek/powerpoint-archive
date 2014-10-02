/**
 * Loops through a single PowerPoint file and exports all the slides
 * as PNG and PPT to <scriptpath>/TEMP.
 *
 * Usage:
 *   cscript dump.js "\\\\some-folder\\some-file"
 *
 * PP reference: http://msdn.microsoft.com/en-us/library/ff743835(v=office.15).aspx
 */
 
function exportSingleSlides(outfolder, infile){
	
	var fso = new ActiveXObject('Scripting.FileSystemObject');
	var fsoForAppend = 8;
	var logfile = fso.CreateTextFile('log.txt', fsoForAppend);
	var app = new ActiveXObject('PowerPoint.Application');
	app.Visible = true;
	
	// Open the presentation
	var presentation = app.Presentations.Open(infile, true); // If not found, script will abort
	var tmpPresentation; // Used to store a new presentation for each slide.

	logfile.WriteLine('Reading ' + infile); 
	WScript.Echo('Reading ' + infile);

	var e = new Enumerator(presentation.Slides);
	var slide;
	e.moveFirst();
	var i = 0;
	var exported = 0;
	var total = presentation.Slides.Count;
	while (!e.atEnd()) {
		i++;
		slide = e.item(); // gets this slide

		// Variable for the unix timestamp. Used in file names below
		var unix = Math.round(+new Date()/1000);

		// Export slide to png
		slide.Export(outfolder + 'slide' + unix + '.png', 'PNG', 1920, 1080);

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
		tmpPresentation.SaveAs(outfolder + 'slide' + unix, 1);
		tmpPresentation.Close();		
		exported++;

		WScript.Echo('Exported ' + exported + ' of ' + total + ' slides'); 

		e.moveNext();
	}

	
	logfile.WriteLine('Exported ' + exported + ' of ' + total + ' slides.'); 
	// Close the presentation
	presentation.Close();
	app.Quit();
}


var selfdir = WScript.ScriptFullName.replace(WScript.ScriptName, '');
exportSingleSlides(selfdir + 'TEMP\\', WScript.arguments(0));
