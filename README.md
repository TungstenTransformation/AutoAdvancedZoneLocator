# AutoAdvancedZoneLocator
This automatically configures zones and OCR profiles in an Advanced Zone Locator  
It can be used to convert Kofax Capture projcets automatically to Kofax Transformation projects  
It can be used to convert a PDF forms document to a Kofax Transformation Project

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
