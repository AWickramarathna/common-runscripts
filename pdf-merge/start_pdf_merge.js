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
//initialize logger
var library = new Object;
eval((new ActiveXObject("Scripting.FileSystemObject")).OpenTextFile("C:/ProgramData/Typefi/typefi-scripts/log.js", 1).ReadAll());


var dumpmetadata = function(arg0, arg1, log, external) {	
	this.wshShell = new ActiveXObject("WScript.Shell");
	var pdfsXmlFilename = arg1;
	var pdf_dir = arg0;
	var external_dir = "";
	
	// setting path / log
	if(external == 'local-graphics'){
		external_dir = pdf_dir;
		log.info("Loading from local-graphics ("+pdf_dir+").");
	}else if(external == 'external-graphics')
		log.info("Loading from external-graphics.");
	else{
		log.info("Load path parameter is empty/ not setted properly");
		log.info("Loading default from external-graphics.");
	}
		
	var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
	xmlDoc.async = false;
	xmlDoc.load(pdfsXmlFilename);
	
	var pdfList = xmlDoc.getElementsByTagName("tps:context[@type='Media']/tps:image");
	log.info("Processing "+pdfList.length+" PDF files");
	
	var myFileSysObj = new ActiveXObject("Scripting.FileSystemObject");	
	
	for (var i = 0; i < pdfList.length; i++) {
	
		var pdfFile = pdfList.item(i).getAttribute("ref");
		var infoFile = pdfFile.replace (/.*[\\\/]/, '').replace(/pdf$/i, 'info');	
		
		if(myFileSysObj.FileExists(external_dir + pdfFile)){
           log.info("File Exists ("+ external_dir + pdfFile+")");
  
			var cmd = "pdftk " + "\""+ external_dir + pdfFile + "\"" + " dump_data output "+ "\""+ pdf_dir + infoFile+"\"" ;
			
			log.info("  CMD: "+cmd);
			var execObj = this.wshShell.Exec(cmd);
			while (execObj.Status == 0) {
				// sit & spin
				WScript.sleep(100);
			}
			
			
			
			WScript.sleep(200);
			log.info("Checking bookmarks...");
			
			//read info file and check for bookmark
			var myInputTextStream = myFileSysObj.OpenTextFile(pdf_dir + infoFile, 1, true);
			var checkTitle = false;
			var checkPageNo = false;

			while(!myInputTextStream.AtEndOfStream){
			  var textLine = myInputTextStream.ReadLine();
			  log.info("  data: "+textLine);
			  
			  //BookmarkTitle: PdftkEmptyString
			  if(textLine === "BookmarkTitle: PdftkEmptyString")
				  checkTitle = true;
				
			  //BookmarkPageNumber: 0	
			  if(textLine === "BookmarkPageNumber: 0")
				  checkPageNo = true;
			  
			  //if title and page is not correct terminate and fail the action
			  if(checkTitle && checkPageNo){
				log.error("Action terminated (Empty bookmark found).");
				log.cancelled();
				WScript.Quit(); 
			  } 		  
			}

			myInputTextStream.Close();
		
		}else
		   log.error("File does not Exists ("+ external_dir + pdfFile+")");
		
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
		log.info('Running PDF Merge Start');
		
		var inputFilePath=json.action.inputs[0].value;
        log.info("inputFilePath : "+inputFilePath);
		var inputFileName=inputFilePath.substring(inputFilePath.lastIndexOf("\\") + 1);
        log.info("inputFileName : "+inputFileName);

		var contentFile = inputFileName;
		var jobFolder = json.environment.jobFolder+"/";
		
		var outputFilePath=json.action.inputs[2].value;
        log.info("outputFilePath : "+outputFilePath);
		var outputFileName=outputFilePath.substring(outputFilePath.lastIndexOf("\\") + 1);
		log.info("outputFileName : "+outputFileName);

		var pdfMergeXMLFile = jobFolder+outputFileName;
		
		//run dump metadata with pdftk
		var graphics_local_external = json.action.inputs[3].value;
        log.info("Parameters : "+json.action.inputs[3].value);

		var pdfMergeXMLFile = jobFolder+outputFileName;
		dumpmetadata(jobFolder, contentFile, log, graphics_local_external);

		
/*		log.info('environment.jobId (' + json.environment.jobId + ').');
		log.info('environment.jobFolder (' + json.environment.jobFolder + ').');
		log.info('environment.systemType (' + json.environment.systemType + ').');
		log.info('environment.locale (' + json.environment.locale + ').');
		log.info('environment.actionId (' + json.environment.actionId + ').');
		log.info('environment.actionLogFile (' + json.environment.actionLogFile + ').');
		log.info('environment.engineVersion (' + json.environment.engineVersion + ').');
		log.info('environment.host (' + json.environment.host + ').');
		log.info('environment.port (' + json.environment.port + ').');
		log.info('environment.inDesignFilestorePath (' + json.environment.inDesignFilestorePath + ').');
		log.info('environment.tpsFilestoreLocation (' + json.environment.tpsFilestoreLocation + ').');
		log.info('action.type (' + json.action.type + ').');
		log.info('action.uiVersion (' + json.action.uiVersion + ').');
		log.info('action.plugin (' + json.action.plugin + ').');
		log.info('action.displayname (' + json.action.displayname + ').');
		log.info('action.description (' + json.action.description + ').');
		log.info('action.pluginversion (' + json.action.pluginversion + ').');
		log.info('action.id (' + json.action.id + ').');
		log.info('action.inputs[0].name (' + json.action.inputs[0].name + ').');
		log.info('action.inputs[0].value (' + json.action.inputs[0].value + ').');
		log.info('action.inputs[1].name (' + json.action.inputs[1].name + ').');
		log.info('action.inputs[1].value (' + json.action.inputs[1].value + ').');
		log.info('action.inputs[2].name (' + json.action.inputs[2].name + ').');
		log.info('action.inputs[2].value (' + json.action.inputs[2].value + ').');
		log.info('action.disabled (' + json.action.disabled + ').'); */
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