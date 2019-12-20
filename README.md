# AutoAdvancedZoneLocator
This automatically configures zones and OCR profiles in an Advanced Zone Locator  
It can be used to convert *Kofax Capture* projects automatically to Kofax Transformation projects  
It can be used to convert a *PDF form* to a Kofax Transformation Project

# Kofax Capture to Kofax Transformation
-	Export your KC project from Kofax Capture as a cab file
- Unzip the Cab File so that admin.xml and admin.dtd are available
- Load „GenerateForms.fpr“ in Kofax Transformation Project Builder
- Call KC2KTM(KCProjectPath As String, KTMProjectsPath As String) with the path to where the KC files are and the path where to create the KTM projects.
- A project will be created for each batch class in the CAB file.
- Load ANY files into KTM
- Press these two icons (“Test Runtime scripts”) & (“Batch View”) so that they are highlighted
![image](https://user-images.githubusercontent.com/47416964/71278567-3e726500-2358-11ea-837f-4118ec35e805.png)
-	Configure the Test Runtime Scripts to call Batch_Open  
![image](https://user-images.githubusercontent.com/47416964/71278600-4fbb7180-2358-11ea-81da-117f8ca51128.png)
- Configure Batch_Open event at the top of the Project script to call KC2KTM(KCInPath, KTMOutPath)
- Press CTRL-F11
- Your project will be created and launched in a new Project Builder.

# PDF Form to Kofax Transformation
This converts a PDF document to an XML file that you can use inside
- download and install git (https://git-scm.com/downloads)
- download and install node.js (https://nodejs.org/en/)
- download and install pdf2json (https://github.com/modesty/pdf2json)
- at the commandline type the following
```
mkdir pdf2json  
cd pdf2json  
git clone git@github.com:modesty/pdf2json.git  
npm install lodash  
npm install asynch  
npm install xmldom  
npm install optimist  
npm install  
"c:\Program Files\nodejs\node.exe" pdf2json.js -f "PDFFilename.pdf" -o "."
```
- you now have a JSON rendition of the PDF document.
- convert it to XML with https://www.freeformatter.com/json-to-xml-converter.html#ad-output
- This XML document is not yet in a format ready for the project. you will need to adjust the xpath expressions in the KT project.
- The Advanced Zone Locator works in millimeters. See https://stackoverflow.com/questions/42494394/pdf2json-page-unit-what-is-it for units conversion.
- See https://github.com/modesty/pdf2json for description of the elements
## Implemented Features
- Import document classes and form types
- Create fields and mappings from locators
- Create Advanced Zone locator, reference document, OCR zones
- Create OCR/OMR profiles, Finereader and Recostar
- Attach ImageCleanupProfile to OCR/OMR Zone
- Barcode Zones
## Not Implemented Features
- Image Cleanup Profile Values
- Create OMR group Zones
- Map All languages used in Recoprofiles (only done Germany, Italy, USA)
- Clean up code. Make consistent, comment well
- Add KC sample files as Layout Classification Samples
