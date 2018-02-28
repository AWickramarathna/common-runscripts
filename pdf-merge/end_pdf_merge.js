/*
 * Copyright 2001-2015 Typefi Systems Pty Ltd. All rights reserved.
 *
 * This is (Windows only) JScript to demonstrate:
 * 1. How to read all the parameters from standard input, and
 * 2. How to log to the action log file.
 *
 * Dependencies:
 * https://svn.typefi.com:9880/Typefi/TPS/trunk/server-plugin-runscript/examples/jscript/src/Log.js
 *
 * Concat dependencies together with this file with a command like this:
 * > type Log.js RunScriptSample.js > C:\ProgramData\Typefi\typefi-scripts\run-script-sample.js
 *
 * Then run it as a Typefi "Run Script" Action like this:
 *
 *   Input [                              ]
 *  Output [                              ]
 * Command [ cscript run-script-sample.js ]
 *
 */
 
//initialize logger
var library = new Object;
eval((new ActiveXObject("Scripting.FileSystemObject")).OpenTextFile("C:/ProgramData/Typefi/typefi-scripts/log.js", 1).ReadAll());

function pdfMerge(arg0, arg1, arg2 ,  log) {
	//	WScript.Echo("List PDF filenames in XML file.");

	var outputFile = "merged.pdf";
	var outputDir = "C:\\temp\\";

	outputDir = arg1;
	outputFile = arg2;
	
	this.wshShell = new ActiveXObject("WScript.Shell");
	var pdfsXmlFilename = arg0;
	var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
	xmlDoc.async = false;
	xmlDoc.load(pdfsXmlFilename);
	
	
	
	// Merge PDFs
	var pdfList = xmlDoc.getElementsByTagName("pdf");
	log.info("Processing "+pdfList.length+" PDF files");
    
	// Main loop for merging PDF files
	for (var i = 0; i < pdfList.length; i++) {
	
	    var tmpOutPDF = outputDir+"tmpPDF"+i+".pdf";
		var pageRange ="";
		
		if (pdfList.item(i).getAttribute("start-page")=="" | pdfList.item(i).getAttribute("end-page") ==""){
			log.info("ERROR:  missing start or end page nubmers for merge in /merge/pdf["+(i+1)+"]");
			
		}
		pageRange = pdfList.item(i).getAttribute("start-page")+"-"+pdfList.item(i).getAttribute("end-page");
		
		var validate_fso = new ActiveXObject("Scripting.FileSystemObject");
		if (!(validate_fso.FileExists(pdfList.item(i).text))) {
			log.info("ERROR: missing PDF file for merge in /merge/pdf["+(i+1)+"]");
		}
		
		
		if ( i == 0) {
			var cmd = "pdftk A=\"" + pdfList.item(i).text + "\" cat A" + pageRange +" output \""+ tmpOutPDF + "\" ";
		} else {
			var tmpInPDF = outputDir+"tmpPDF"+(i-1)+".pdf";
			var cmd = "pdftk A=\"" + tmpInPDF + "\" B=\"" + pdfList.item(i).text + "\" cat A B"+ pageRange+" output \""+ tmpOutPDF+"\"";
		}	
		log.info("  CMD: "+cmd);
		var execObj = this.wshShell.Exec(cmd);
		while (execObj.Status == 0) {
			// sit & spin
			WScript.sleep(100);
		}
	  
	  
	  if (i == pdfList.length-1) {
	    
		var finalPDFInfoFileName = outputDir+"merge.info";
		var tmpInfoFileName = outputDir+"tmp.info";
		var bookMarksFileName= outputDir+"merge_bookmarks.info";
		
	    log.info("tmp File:"+tmpOutPDF);
		log.info("output File:"+outputFile);
		
		//fso.moveFile(tmpOutPDF,outputDir+outputFile);
		log.info("Merging Metadata");
		//Merge in mata-data files
		
		
		//Read bookmarks file
		var bm_fso = new ActiveXObject("Scripting.FileSystemObject");
		var bookMarksFile = bm_fso.OpenTextFile(bookMarksFileName,1,true);
		
		if (bookMarksFile.AtEndOfStream)
			var bookMarks = " ";
        else
			var bookMarks = bookMarksFile.ReadAll();
		bookMarksFile.Close();

		//extract the other PDF meta-data from the output file
		var cmd = "pdftk " + "\"" + tmpOutPDF + "\"" + " dump_data output "+ "\"" +tmpInfoFileName+ "\"";
		log.info("  CMD: "+cmd);
		var execObj = this.wshShell.Exec(cmd);
		while (execObj.Status == 0) {
			// sit & spin
			WScript.sleep(1000);
		}
		var tmp_fso = new ActiveXObject("Scripting.FileSystemObject");
		var tmpInfoFile = tmp_fso.OpenTextFile(tmpInfoFileName,1,true);
		if (tmpInfoFile.AtEndOfStream)
			var tmpInfo=" ";
		else
			var tmpInfo=tmpInfoFile.ReadAll();;
		tmpInfoFile.Close();
		
		var ff_fso = new ActiveXObject("Scripting.FileSystemObject");
		var finalPDFInfoFile = ff_fso.OpenTextFile(finalPDFInfoFileName,2,true);
		finalPDFInfoFile.Write(tmpInfo);
		finalPDFInfoFile.Write(bookMarks);
		finalPDFInfoFile.Close();
		
		//merge the meta-data to the PDFtk
		var cmd = "pdftk " + "\"" + tmpOutPDF + "\"" + " update_info "+"\""+finalPDFInfoFileName+"\""+" output "+ "\""+outputFile+"\"";
		log.info("  CMD: "+cmd);
		var execObj = this.wshShell.Exec(cmd);
		while (execObj.Status == 0) {
			// sit & spin
			WScript.sleep(100);
		}
		
		
		//clean up
		
		var fso = new ActiveXObject("Scripting.FileSystemObject");
		log.info("Cleaning up...");
		for (var j = 0; j < (pdfList.length); j++) {
		  var tmpClnPDF = outputDir+"tmpPDF"+j+".pdf";
		  log.info("Deleting:"+tmpClnPDF);
		  //fso.DeleteFile(tmpClnPDF);
		}
		//fso.DeleteFile(tmpInfoFileName);
		//fso.DeleteFile(bookMarksFileName);
		//fso.DeleteFile(finalPDFInfoFileName);
	  }
    }
	
	// Rotate Pages
	var fileToRotatePages = outputFile;
	var rotateList = xmlDoc.getElementsByTagName("rotate");
	log.info("Rotating a total of "+rotateList.length+" pages");
	
    
	// Main loop for merging PDF files
	for (var r = 0; r < rotateList.length; r++) {
		
		var tmpOutPDF = outputDir+"tmpRPDF"+r+".pdf";
		
		
		if (rotateList.item(r).getAttribute("page")!="") {
		    rotatePage = rotateList.item(r).getAttribute("page");
			if ( r == 0) {
				var cmd = "pdftk \"" + fileToRotatePages + "\" rotate " + rotatePage + "left output \"" + tmpOutPDF + "\"";
			} else {
				var tmpInPDF = outputDir+"tmpRPDF"+(r-1)+".pdf";
				var cmd = "pdftk \"" + tmpInPDF + "\" rotate " + rotatePage + "left output \"" + tmpOutPDF + "\"";
			}
			log.info(cmd);
			var execObj = this.wshShell.Exec(cmd);
			while (execObj.Status == 0) {
				// sit & spin
				WScript.sleep(100);
			}	
		}
		else
		{
			log.info("ERROR: missing page nubmer for rotation in /merge/rotate["+(r+1)+"]");
		}
		//clean up
		
		if (r == rotateList.length-1) {
		
			var fso = new ActiveXObject("Scripting.FileSystemObject");
			fso.CopyFile (tmpOutPDF,fileToRotatePages);
			log.info("Cleaning up...");
			for (var j = 0; j < (rotateList.length); j++) {
				var tmpClnPDF = outputDir+"tmpRPDF"+j+".pdf";
				log.info("Deleting:"+tmpClnPDF);
				//fso.DeleteFile(tmpClnPDF);
			}
		}
	
	}

}; 
 
(function() {
	try {
		// 1. read standard input into a string
		var stdin = WScript.StdIn.ReadAll();

		// 2. convert the string to a JavaScript object (TODO: eval() is unsafe)
		var json = eval('(' + stdin + ')');
		    
		// 3. open the action log file and start writing
		var log = new Log(json.environment.actionLogFile);
		log.commenced();
		
		
        var inputFilePath=json.action.inputs[0].value;
        log.info("inputFilePath : "+inputFilePath);
		var inputFileName=inputFilePath.substring(inputFilePath.lastIndexOf("\\") + 1);
		log.info("inputFileName : "+inputFileName);

	    var contentFile = json.environment.jobFolder+inputFileName;
		var jobFolder = json.environment.jobFolder+"/";	
        log.info("inputFileName : "+inputFileName);	
		
        log.info("outputFile : "+json.action.inputs[1].value);
		var labelJob = json.action.inputs[1].value;

		var pdfMergeXMLFile = jobFolder+'merge.xml';
			
		
		pdfMerge(pdfMergeXMLFile, jobFolder, labelJob + ".pdf" ,  log) 
			/*
		var cmd = "cscript //nologo C:/ProgramData/Typefi/typefi-scripts/pdfMerge/pdf-merge.js \"" + pdfMergeXMLFile + "\" \"" + jobFolder +"\" \""+ labelJob + ".pdf\"";
		log.info("Executing: "+cmd);
		
		var wshShell = new ActiveXObject("WScript.Shell");
		var execObj = wshShell.Exec(cmd);
		log.info("After Exec call");
		while (execObj.Status == 0) {
			// sit & spin
			 if (!execObj.StdOut.AtEndOfStream) {
				  log.info(execObj.StdOut.ReadLine());
			 }
			WScript.sleep(100);
		}	
		if (execObj.ExitCode == 0) {
			// things went ok
			log.info("After exec");
		}
		else {
			throw "Error executing ant.  Exit Code: " + execObj.ExitCode;
		}
		*/
		
		log.completed();
	} catch (e) {
		log.error(e.message);
		log.cancelled();
	} finally {
		log.close();
	}
})();
/* example data from standard input:
{
	"environment":{
		"jobId": 2147478384,
		"jobFolder": "d:/Typefi/v8_filestore/Lincoln/workflows/run-script/2015-09-08 09.46.40",
		"systemType": "DESKTOP",
		"locale": "en-US",
		"actionId": "2147478384-1",
		"actionLogFile": "d:/Typefi/v8_filestore/Lincoln/workflows/run-script/2015-09-08 09.46.40/logs/01 Run Script.log",
		"engineVersion": "",
		"host": "",
		"port": "",
		"inDesignFilestorePath": "",
		"tpsFilestoreLocation": ""
		},
	"action":{
		"type": "runscript",
		"uiVersion": 1,
		"plugin": "typefi-plugin-runscript",
		"displayname": "Run Script",
		"description": "Run Script",
		"pluginversion": "?",
		"id": "1",
		"inputs": [
			{
				"name": "input",
				"value": "d:\\Typefi\\v8_filestore\\Lincoln\\Content\\input.cxml"
			},
			{
				"name": "cmd",
				"value": "cscript test.js //NoLogo"
			},
			{
				"name": "output",
				"value": "d:\\Typefi\\v8_filestore\\Lincoln\\Content\\output.cxml"
			}
		],
		"disabled": false
	}
}
*/