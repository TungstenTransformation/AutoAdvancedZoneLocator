'#Reference {50A7E9B0-70EF-11D1-B75A-00A0C90564FE}#1.0#0#C:\Windows\SysWOW64\shell32.dll#Microsoft Shell Controls And Automation#Shell32
'#Reference {F5078F18-C551-11D3-89B9-0000F81FE221}#6.0#0#C:\Windows\System32\msxml6.dll#Microsoft XML, v6.0#MSXML2
'#Reference {420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\Windows\SysWOW64\scrrun.dll#Microsoft Scripting Runtime#Scripting
'#Reference {32AC4EE4-094A-4225-8F82-51277730B675}#4.0#0#C:\Program Files (x86)\Common Files\Kofax\Components\CscForms2.dll#Kofax Cascade Forms Processing 4.0#CSCFORMSLib
'#Reference {4DE5DB81-4963-4D5C-8D0A-A3009CC031E2}#4.0#0#C:\Program Files (x86)\Common Files\Kofax\Components\CscAdvZoneLoc2.dll#Kofax Cascade Advanced Zone Locator 4.0#CscAdvZoneLocLib
'#Reference {CE882FA9-3723-4819-B45D-9DA565487491}#5.5#0#C:\Program Files (x86)\Common Files\Kofax\Components\DatabaseDialog.tlb#Kofax Transformation database lookup for script integration#DatabaseDialog
'#Language "WWB-COM"
Option Explicit
'TODO
' Copy recognition profiles
' Copy anchors

Private Sub Batch_Open(ByVal pXRootFolder As CASCADELib.CscXFolder)
   KC2KTM("C:\Users\david.wright\Documents\projects\KC2KTM\LESSON_4","C:\Users\david.wright\Documents\projects\KC2KTM\out")
End Sub

Private Sub Document_AfterClassifyXDoc(ByVal pXDoc As CASCADELib.CscXDocument)
   KC2KTM("C:\Users\david.wright\Documents\projects\KC2KTM\LESSON_4","C:\Users\david.wright\Documents\projects\KC2KTM\out")
End Sub

Private Sub KC2KTM(KCProjectPath As String, KTMProjectsPath As String)
   'Creates a new KTM project for each batch class in the exported cab file
   'Each KC documentClass becomes a KTM class
   'Each KC formType becomes a subclass to a documentClass
   'Each KC IndexZone becomes a zone in an advanced zone locator
   Dim xml As MSXML2.DOMDocument60
   Set xml = New MSXML2.DOMDocument60
   xml.setProperty("ProhibitDTD",False)
   xml.resolveExternals=False
   Dim success As Boolean
   Dim nodeList As MSXML2.IXMLDOMNodeList
   Dim Batch As MSXML2.IXMLDOMNode
   Dim docClass As MSXML2.IXMLDOMNode
   Dim formType As MSXML2.IXMLDOMNode
   Dim recoProfile As MSXML2.IXMLDOMNode
   Dim field As MSXML2.IXMLDOMNode
   Dim batchName,docClassName,formTypeName,projectFileName,projectPath As String
   Dim AZL As CscLocatorDef' CscAdvZoneLocator
   success = xml.Load(KCProjectPath & "\admin.xml")
   If success = False Then MsgBox xml.parseError.reason: Exit Sub
   Dim pr As CscProject
   For Each Batch In xml.selectNodes("//BatchClasses/BatchClass")
      batchName=Batch.Attributes.getNamedItem("Name").Text
      Set pr = New CscProject
      While pr.RecogProfiles.Count>0
'         pr.RecogProfiles.Remove(0)
      Wend
      projectPath=KTMProjectsPath & "\" & batchName
      projectFileName=projectPath & "\" & batchName & ".fpr"
      Dim fso As New Scripting.FileSystemObject
      If fso.FolderExists(projectPath) = True Then fso.DeleteFolder(projectPath,True)
      fso.CreateFolder(projectPath)
      pr.Save(projectFileName)
      For Each recoProfile In xml.selectNodes("//RecognitionProfiles/RecognitionProfile")
         Project_AddRecoProfile(pr,recoProfile)
      Next
      For Each docClass In Batch.selectNodes("DocumentClassLinks/DocumentClassLink/@Name")
         pr.AddClass(docClass.Text,"")
         For Each field In xml.selectNodes("//DocumentClasses/DocumentClass[@Name='" & docClass.Text & "']/IndexFieldDefinitions/IndexFieldDefinition")
            ProjectClass_AddField(pr.ClassByName(docClass.Text),field)
         Next
         For Each formType In xml.selectNodes("//DocumentClasses/DocumentClass[@Name='" & docClass.Text & "']/FormTypes/FormType")
            formTypeName=formType.Attributes.getNamedItem("Name").Text
            pr.AddClass(formTypeName,docClass.Text)
            Set AZL=ProjectClass_AddZoneLocator(Project.ClassByName(formTypeName),"AZL")
            AZL_ImportZonesFromKC(AZL,formType,Project.ClassByName(docClassName))
            'TODO: Map each subfield to field
            'fieldDef.Locator = "AZL"
            'fieldDef.LocatorSubField= fieldDef.Name
            'Assign Advanced Zone Locator to Locator Definition to update Subfields
            'locdef.AssignLocatorMethod(AZL)
         Next
      Next
      pr.Save(pr.FileName)
   Next
End Sub

Private Sub AZL_ImportZonesFromKC(AZL As CscLocatorDef, formType As MSXML2.IXMLDOMNode,docClass As CscClass)
   ' Clean target folder

   ' Attach sample image to AZL
   Dim XDoc As CscXDocument
   Set XDoc = pXRootFolder.DocInfos(0).XDocument
   pr.AddSampleDoc(formType.Attributes.getNamedItem("Name").Text,XDoc)
   Dim zones As MSXML2.IXMLDOMNode
   Dim zone As MSXML2.IXMLDOMNode
   'Create OCR/OMR Zones
   Set zones= xml.selectNodes("//FormType[@Name='" & formName & "']/IndexZones/IndexZone")
   For Each zone In nodeList
      Dim l,t,w,h,p As Long
      Dim n As String
      l=CLng(node.Attributes.getNamedItem("ZoneLeft").Text)
      t=CLng(node.Attributes.getNamedItem("ZoneTop").Text)
      w=CLng(node.Attributes.getNamedItem("ZoneWidth").Text)
      h=CLng(node.Attributes.getNamedItem("ZoneHeight").Text)
      p=CLng(node.Attributes.getNamedItem("PageNumber").Text)-1
      n=Replace(node.Attributes.getNamedItem("IndexFieldDefinitionName").Text," ","_")
      AZL_addOMRZone(AZL,n,l,t,w,h,p,AZL.Zones.GetNextId)
   Next

   ' Add group zone
   Dim groupId As Integer
   groupId=AZL_addOMRGroup(AZL, "OMRGroup",36.0,92.0,50.0,7.5,0)

   ' Add OMR zones
   'AZL_addOMRGroup(AZL,"OMRZone0",39.0,93.0,4.5,5.5,0,groupId)
   'AZL_addOMRGroup(AZL,"OMRZone1",47.6,93.0,4.5,5.5,0,groupId)
   'AZL_addOMRGroup(AZL,"OMRZone2",56.2,93.0,4.5,5.5,0,groupId)
   'AZL_addOMRGroup(AZL,"OMRZone3",64.8,93.0,4.5,5.5,0,groupId)

   ' Anchor
   AZL_addAnchor(XDoc, AZL,"topleft",0.0,0.0,50.0,50.0,0,CscAdvZoneLocReferencePointTypeTopLeft)
   mapAllAnchorsToAllZones(AZL)
End Sub

Function AZL_addOMRGroup(AZL As CscAdvZoneLocator, Name As String, Left As Double, Top As Double, Width As Double, Height As Double, pageNr As Integer) As Integer
   Dim zone As New CscAdvZoneLocZone
   zone.ID=AZL.Zones.GetNextId()
   zone.Name=Name
   zone.Left=Left
   zone.Top=Top
   zone.Width=Width
   zone.Height=Height
   zone.PageNr=pageNr
   zone.GroupId=zone.ID
   zone.RecogProfileId=Project.RecogProfiles.DefaultProfileIdZrOmr
   AZL.Zones.Append(zone)
   ' Make the subfield and connect it to the zone
   AZL_addSubfieldAndMapping(AZL, zone)
   ' Return value
   AZL_addOMRGroup=zone.ID
End Function

Sub AZL_addSubfieldAndMapping(AZL As CscAdvZoneLocator, Zone As CscAdvZoneLocZone)
   Dim subfield As New CscAdvZoneLocSubfield
   subfield.ID=AZL.SubFields.GetNextId
   subfield.Name=Zone.Name
   subfield.ResultType=CscAdvZoneLocSubfieldResultTypeBest
   AZL.SubFields.Append(subfield)
   Dim mapping As New CscAdvZoneLocMapping
   mapping.SubfieldId=subfield.ID
   mapping.ZoneId=Zone.ID
   AZL.Mappings.Append(mapping)
End Sub

Sub AZL_addOMRZone(AZL As CscAdvZoneLocator, Name As String, Left As Double, Top As Double, Width As Double, Height As Double, pageNr As Integer, groupId As Integer)
   Dim Zone As New CscAdvZoneLocZone
   Zone.ID=AZL.Zones.GetNextId
   Zone.Name=Name
   Zone.Left=Left
   Zone.Top=Top
   Zone.Width=Width
   Zone.Height=Height

   Zone.PageNr=pageNr
   Zone.GroupId=groupId
   Zone.RecogProfileId=Project.RecogProfiles.DefaultProfileIdZrOmr
   AZL.Zones.Append(Zone)
   ' Make the subfield and connect it to the zone
   AZL_addSubfieldAndMapping(AZL, Zone)
End Sub

Sub AZL_addAnchor(XDoc As CscXDocument, AZL As CscAdvZoneLocator, Name As String, Left As Double, Top As Double, Width As Double, Height As Double, pageNr As Integer, refPointType As CscAdvZoneLocReferencePointType)
   Dim anchor As New CscAdvZoneLocAnchorZone
   anchor.ID=AZL.AnchorZones.GetNextId()
   anchor.Name=Name
   anchor.Left=Left
   anchor.Top=Top
   anchor.Width=Width
   anchor.Height=Height
   anchor.Page=pageNr
   anchor.CanTouchBorder=False
   anchor.MaxAnchorWidth=999.9 'mm
   anchor.MaxAnchorHeight=999.9 'mm
   anchor.MinAnchorWidth=5.0 'mm
   anchor.MinAnchorHeight=5.0 'mm
   anchor.Threshold=0.70
   anchor.ReferencePointType=refPointType
   anchor.ResetInvalid()
   findAnchor(XDoc, anchor)
   AZL.AnchorZones.Append(anchor)
End Sub

Sub findAnchor(XDoc As CscXDocument, anchor As CscAdvZoneLocAnchorZone)
      Dim Image As CscImage
      Set Image = XDoc.CDoc.Pages(anchor.Page).GetImage()

      ' Find line pattern
      Dim Finder As New CscLinePatternFinder

      Finder.CanTouchBorder = anchor.CanTouchBorder
      Finder.MaxHeightMM = anchor.MaxAnchorHeight
      Finder.MaxWidthMM = anchor.MaxAnchorWidth
      Finder.MinHeightMM = anchor.MinAnchorHeight
      Finder.MinWidthMM = anchor.MinAnchorWidth

      ' Convert mm to pixel
      Dim Left, Top, Width, Height As Integer
      Left = Round(MmToPixel(anchor.Left, Image.XResolution))
      Top = Round(MmToPixel(anchor.Top, Image.YResolution))
      Width = Round(MmToPixel(anchor.Width, Image.XResolution))
      Height = Round(MmToPixel(anchor.Height, Image.YResolution))

      ' Map types
      Dim t As CscLinePatternType
      Select Case (anchor.ReferencePointType)
         Case CscAdvZoneLocReferencePointType.CscAdvZoneLocReferencePointTypeTopLeft
            t = CscLinePatternType.CscPTCornerUpperLeft
         Case CscAdvZoneLocReferencePointType.CscAdvZoneLocReferencePointTypeTopRight
            t = CscLinePatternType.CscPTCornerUpperRight
         Case CscAdvZoneLocReferencePointType.CscAdvZoneLocReferencePointTypeBottomLeft
            t = CscLinePatternType.CscPTCornerLowerLeft
         Case CscAdvZoneLocReferencePointType.CscAdvZoneLocReferencePointTypeBottomRight
            t = CscLinePatternType.CscPTCornerLowerRight
         Case Else
            t = CscLinePatternType.CscPTNone
      End Select

      ' Find anchor pattern
      Finder.Find(t, Image, Left, Top, Width, Height)

      If Finder.LinePatternCount > 0 Then
         Dim lpi As CscLinePatternInfo
         Set lpi = Finder.GetLinePattern(0)

         ' Map types back
         Select Case (lpi.LinePatternType)
            Case CscLinePatternType.CscPTCornerUpperLeft
               anchor.AnchorType = CscAdvZoneLocAnchorType.CscAdvZoneLocAnchorTypeCornerUpperLeft
            Case CscLinePatternType.CscPTCornerUpperRight
               anchor.AnchorType = CscAdvZoneLocAnchorType.CscAdvZoneLocAnchorTypeCornerUpperRight
            Case CscLinePatternType.CscPTCornerLowerLeft
               anchor.AnchorType = CscAdvZoneLocAnchorType.CscAdvZoneLocAnchorTypeCornerLowerLeft
            Case CscLinePatternType.CscPTCornerLowerRight
               anchor.AnchorType = CscAdvZoneLocAnchorType.CscAdvZoneLocAnchorTypeCornerLowerRight
            Case Else
               anchor.AnchorType = CscAdvZoneLocAnchorType.CscAdvZoneLocAnchorTypeNone
         End Select

         anchor.AnchorLeft = PixelToMm(lpi.StartX, Image.XResolution)
         anchor.AnchorTop = PixelToMm(lpi.StartY, Image.YResolution)
         anchor.AnchorWidth = PixelToMm(lpi.Width, Image.XResolution)
         anchor.AnchorHeight = PixelToMm(lpi.Height, Image.YResolution)
         anchor.AnchorCenterX = PixelToMm(lpi.CenterX, Image.XResolution)
         anchor.AnchorCenterY = PixelToMm(lpi.CenterY, Image.YResolution)
      End If
End Sub

Function MmToPixel(Mm As Double, Resolution As Double) As Double
   MmToPixel = (Mm / 25.4 * Resolution)
End Function

Function PixelToMm(Pixel As Long, Resolution As Double) As Double
   PixelToMm = (Pixel * 25.4 / Resolution)
End Function

Function Round(Value As Double) As Integer
   If Value < 0 Then
      Round = Value - 0.5
   Else
      Round = Value + 0.5
   End If
End Function

Sub mapAllAnchorsToAllZones(AZL As CscAdvZoneLocator)
   Dim a As Integer
   Dim z As Integer
   For a = 0 To AZL.AnchorZones.Count-1
      For z = 0 To AZL.Zones.Count-1
         Dim IsGroup As Boolean
         IsGroup = (AZL.Zones(z).ID = AZL.Zones(z).GroupId)

         If Not IsGroup Then
            Dim am As New CscAdvZoneLocAnchorMapping
            am.AnchorZoneId=AZL.AnchorZones(a).ID
            am.ZoneId=AZL.Zones(z).ID
            AZL.AnchorMappings.Append(am)
            AZL.Zones(z).ValidAnchorCount=-1 'We don't require (but we would like them all!) any anchors to match for registration
         End If
      Next
   Next
End Sub


Private Sub Project_AddRecoProfile(pr As CscProject,KCrecoProfile As MSXML2.IXMLDOMNode)
   Dim recoProfile As IMpsRecogProfile
   'todo
End Sub


Private Sub ProjectClass_AddField(cl As CscClass, KCfield As MSXML2.IXMLDOMNode)
   'Adds a KC index field to a KTM document class
   Dim fieldDef As New CscFieldDef
   fieldDef.Name = Replace(KCfield.Attributes.getNamedItem("Name").Text," ", "_")
   fieldDef.FieldType = CscExtractionFieldType.CscFieldTypeSimpleField
   cl.Fields.Add(fieldDef)
   fieldDef.AlwaysValid=(KCfield.Attributes.getNamedItem("Hidden").Text=1)
   Return fieldDef
End Sub

Private Function ProjectClass_AddZoneLocator(cl As CscClass,azlName As String) As CscLocatorDef
   'Adds an empty Advanced Zone Locator to a KTM class
   Dim locdef As New CscLocatorDef
   locdef.AssignLocatorMethod(New CscAdvZoneLocator)
   locdef.Name=azlName
   cl.Locators.Add(locdef)
   With locdef.LocatorMethod
      .RegMetaMode = CscRegMetaType.CscRegMetaTypeNone
      .RegModes = CscRegType.CscRegTypeNone
   End With
   Return locdef
End Function
