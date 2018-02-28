
Follow these steps to install and configure the PDF Merge Utility:

1. Download and install PDFtk Server from https://www.pdflabs.com/tools/pdftk-server/. Reboot the desktop computer or server to ensure complete integration with Typefi Publish.
2. Test PDFtk Server using the Windows command prompt. Type pdftk --version and press the enter key. A short message that includes the version number should appear.
3. Install the PDF Merge Utility by copying the following files into the Typefi Server scripts directory:
	- start_pdf_merge.js
	- end_pdf_merge.js
4. On the Typefi Server, copy the following command into the job.start event within the job option settings:
	- cscript start_pdf_merge.js //NoLogo
5. On the Typefi Server, copy the following command into the job.end event within the job option settings:
	- cscript end_pdf_merge.js //NoLogo

Assumptions:

* The Typefi project has been configured to produce the front-cover, padding, and back-cover PDF files, and they are present in the filestore.
* Running the job produces the merge.xml file.
* Absolute File Paths must be enabled.

Additional information:

* PDFtk Server: http://www.pdflabs.com/tools/pdftk-server/ 
* PDFtk Server Manual: http://www.pdflabs.com/docs/pdftk-man-page/
* Email sid.steward@pdflabs.com with questions or bug reports regarding PDFtk Server , and include "pdftk" in your email subject.