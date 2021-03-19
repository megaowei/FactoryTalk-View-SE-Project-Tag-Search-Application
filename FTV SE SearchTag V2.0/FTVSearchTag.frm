VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FTVSearchTag 
   Caption         =   "FactoryTalk View Search Tag Dialog"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10140
   OleObjectBlob   =   "FTVSearchTag.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FTVSearchTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'******************************Instruction************************************************************
'*****The Code has been updated at 20190601 to support opening faceplate of tags as navigating display pages. The function only supports PlantPAX 4.0 or latest version.
'*****For HMI tags(like OPC server tag path) , should distribute the path of tag to make it easy to distinguish the name of faceplate
'******************************Instruction************************************************************
'As you use this VBA code for searching function in FTV SE system, you should firstly use XML Merge tool to merge all xml files
'which are exported from FTV SE Studio Display part into one XML file <Search_XML_File.xml>.
'Please note that it is the best to create new GUI that is the same with this example, for example all objects(buttons,Testbox control
'Combobox control) must use the same name,then copy the whole VBA code into VBA IDE.
'*****************************************************************************************************
Dim TagName1 As String
Dim oMyXmlDoc As DOMDocument 'Ceate new XML document
Dim oMyXmlNode As IXMLDOMNode
Dim oMyxmlNodes As Object 'Ceate new XML nodes object
Dim A1, strpath, strTag, str, str1, str2, xlsFile, TagFileStr, TagText, Tagstr, Devicestr, pagestr, strm, strTagPath, tempTagFileStr As String
Dim i, j, NodeNum, n, NumFile, searchj, searchFile, Itemi, Itemj, numItem, TagIndexInDislayArray As Integer
Dim sscPrimaryStatus As gfxServerStatusConstants
Dim sscSecondaryStatus As gfxServerStatusConstants
Dim SFullNameOfHCS As String
Dim ActiveComputerName As String
Dim Tagnames(0 To 2000), ScreenName(0 To 2000), DisplaylistItem(2000, 50) As String ' For DisplaylistItem array, the first column is tagpath,the second column is tagname, the others are the pages that contain the tag
Dim ESDTagFaceplate(100, 2), FGSTagFaceplate(100, 2) As String, FAPTagFaceplate(100, 2) As String  'The two arrary are distributed for OPC/HMI tags at AADvance project

'*****The part is to display the screen that you click at ComboBox list**************************

Private Sub ScreenListBox_Click()

   Dim DisplayName, strParameter, strTagPathInArray, strTagNameInArray As String
   Dim indexTagMark As Integer
   
   Call InitialTagFaceplateArray
   
   indexTagMark = 1
   strTagPathInArray = DisplaylistItem(TagIndexInDislayArray, 1) 'NavToFaceplate #102 #103 "#120" "#121"
   strTagNameInArray = DisplaylistItem(TagIndexInDislayArray, 2)
   DisplayName = ScreenListBox.Value 'Get the screen name that you select

   Application.ExecuteCommand ("Display " & DisplayName) ' send display command to HMI server to display screen
   
   If InStr(strTagPathInArray, "]") <> 0 Then
   
      strParameter = strTagPathInArray & strTagNameInArray & " " & strTagPathInArray
      Application.ExecuteCommand ("NavToFaceplate " & strParameter) ' send display command to HMI server to display faceplate

   ElseIf InStr(strTagPathInArray, "\") <> 0 Then '************** This part is to support HMI/OPC tags (like AADvance project) to open faceplate as navigating pages***************************
   
   '********************************************************************************************
          '******************* if there are many fold in HMI/OPC tag, you can copy this part code, then just change ESD to xxx which you want,*******
          '******************* firstly you should create xxxTagFaceplate string array, and complete the relation logic at Sub InitialTagFaceplateArray()**********************************
          If InStr(strTagPathInArray, "ESD") <> 0 Then
             
             Do While ESDTagFaceplate(indexTagMark, 1) <> ""
             
                If InStr(1, strTagNameInArray, ESDTagFaceplate(indexTagMark, 1)) <> 0 Then
                
                   strParameter = ESDTagFaceplate(indexTagMark, 2) & " /T " & strTagPathInArray & strTagNameInArray
                   
                   Exit Do
                   
                End If

                indexTagMark = indexTagMark + 1
                
             Loop
             indexTagMark = 1
          '****************************************************************
          ElseIf InStr(strTagPathInArray, "FGS") <> 0 Then
          
              Do While FGSTagFaceplate(indexTagMark, 1) <> ""
             
             
                If InStr(1, strTagNameInArray, FGSTagFaceplate(indexTagMark, 1)) <> 0 Then
                
                   strParameter = FGSTagFaceplate(indexTagMark, 2) & " /T " & strTagPathInArray & strTagNameInArray
                   
                   Exit Do
                   
                End If

                indexTagMark = indexTagMark + 1
                
             Loop
             indexTagMark = 1
          ElseIf InStr(strTagPathInArray, "FAP") <> 0 Then
          
              Do While FAPTagFaceplate(indexTagMark, 1) <> ""
             
             
                If InStr(1, strTagNameInArray, FAPTagFaceplate(indexTagMark, 1)) <> 0 Then
                
                   strParameter = FAPTagFaceplate(indexTagMark, 2) & " /T " & strTagPathInArray & strTagNameInArray
                   
                   Exit Do
                   
                End If

                indexTagMark = indexTagMark + 1
                
             Loop
             indexTagMark = 1
          End If
          
   '********************************************************************************************
          
          Application.ExecuteCommand ("Display " & strParameter) ' send display command to HMI server to display faceplate
      
   End If

   Erase DisplaylistItem  ' Clear  string arrays
   Erase Tagnames ' Clear  string arrays
   Erase ScreenName ' Clear  string arrays
   Erase ESDTagFaceplate
   Erase FGSTagFaceplate
   
End Sub

Private Sub Searchbutton_Click()

On Error GoTo errHandler

'*****The purpose of this part is to set XML file in redundancy HMI server at Network Distribute System******

     'SFullNameOfHCS = "/Hull:Hull"  ' "/RootAreaName:HMIServerName"
     'GetServerStatus SFullNameOfHCS, sscPrimaryStatus, sscSecondaryStatus, ActiveComputerName
     'If sscPrimaryStatus = gfxServerStatusActive And ActiveComputerName = "HCSSVRA" Then
     '   strpath = "\\HCSSVRA\Users\Public\Documents\SearchTag\Search_XML_File.xml"
     'End If
     'If sscSecondaryStatus = gfxServerStatusActive And ActiveComputerName = "HCSSVRB" Then
     '   strpath = "\\HCSSVRB\Users\Public\Documents\SearchTag\Search_XML_File.xml"
     'End If
strpath = "C:\Users\Administrator\Desktop\New folder\Search_XML_File.xml"
'************************************************************************************************************
'If FTV SE system is not redundancy system, please delete upside part, you can use command like below:
'  strpath = "****\******\****\SearchTag\Search_XML_File.xml"   the path contains <Search_XML_File.xml> file
'************************************************************************************************************
     
     TagListBox.Clear  ' Clear listbox as you start searching
     ScreenListBox.Clear  ' Clear listbox as you start searching
     Erase DisplaylistItem  ' Clear  string arrays
     Erase Tagnames
     Erase ScreenName
     NumFile = 0   ' Counter of screens that contain tag/device which you want
     Itemi = 1
     Itemj = 1
     numItem = 4  ' this variable is for counting how many screens contain seearched tag that you select
     
     If TagInput.Value = "" Then ' Send a alarm message as you have no type tag name
     
        MsgBox ("Erro...Please input tag name!")
     
     Exit Sub
     
     End If
  
     If InStr(1, TagInput.Value, "-") <> 0 Then   ' this is for replacing "-", it is based on your program
        Tagstr = Replace(TagInput.Value, "-", "_")
     Else
        Tagstr = TagInput.Value
     End If
     
     Tagstr = UCase(Tagstr)   ' upper case the tagstr, this is based on your program, if your tags are low case, you can use LCASE() function
     
     Set oMyXmlDoc = New DOMDocument   ' Create new XML Document object for loading XML file
     
     oMyXmlDoc.async = False
     
     oMyXmlDoc.Load (strpath) 'load XML file <Search_XML_File.xml>
    
     str = "//Tag"

     Set oMyxmlNodes = oMyXmlDoc.selectNodes(str) ' Get all <Tag> elements from the loaded XML file.

     NodeNum = oMyxmlNodes.length ' Get the number of <Tag> elements
     
     searchFile = 1 ' The pointer for delete the repeated items that are contained at ComboBox list
     
     For j = 0 To NodeNum - 1
     
         TagFileStr = oMyxmlNodes.Item(j).Attributes(0).Text ' Get tagname attribute of one element, you can open <Search_XML_File.xml> to see detail info
         
          pagestr = Left(oMyxmlNodes.Item(j).Attributes(1).Text, Len(oMyxmlNodes.Item(j).Attributes(1).Text) - 4) ' Get screenname attribute of one element

         If InStr(1, TagFileStr, Tagstr) <> 0 Then ' decide whether tagname attribute contains tagstr or not

'********************************************************************************************************************
'For this part, you can modify the logic based on different project. The purpose is to get the real device/tag name from Tag name attribute
'********************************************************************************************************************
              If InStr(1, TagFileStr, "]") <> 0 Then
           
                 If InStr(1, TagFileStr, ".") <> 0 Then
                 
                     strTag = Split(Split(TagFileStr, "]")(1), ".")(0) ' split tagname to get what you want
              
                 ElseIf InStr(1, TagFileStr, "}") <> 0 Then
                  
                     strTag = Split(Split(TagFileStr, "]")(1), "}")(0)
                     
                 Else
                 
                     strTag = Split(TagFileStr, "]")(1)
                 
                 End If
                 
                 If InStr(1, TagFileStr, "{") <> 0 Then
                    strTagPath = Split(Split(TagFileStr, "]")(0), "{")(1) & "]"
                 Else
                    strTagPath = Split(TagFileStr, "]")(0) & "]"
                 End If
              
             ElseIf InStr(1, TagFileStr, "\") <> 0 Then
                     If InStr(1, TagFileStr, "}") <> 0 Then
                        tempTagFileStr = Split(Split(TagFileStr, "}")(0), "{")(1)
                        strTag = Mid(tempTagFileStr, InStrRev(tempTagFileStr, "\") + 1) ' If tagname attribute  contain "\", for HMI tags/OPC tags
                        strTagPath = Left(tempTagFileStr, InStrRev(tempTagFileStr, "\")) 'this is one method for spliting string
                     Else
                        strTag = Mid(TagFileStr, InStrRev(TagFileStr, "\") + 1) ' If tagname attribute  contain "\", for HMI tags/OPC tags
                        strTagPath = Left(TagFileStr, InStrRev(TagFileStr, "\")) 'this is one method for spliting string
                     End If
             Else
                 strTag = TagFileStr
                 strTagPath = TagFileStr
  
             End If

             
'********************************************************************************************************************

             For searchj = 1 To searchFile
                
                If searchFile = 1 Then ' Add the first finded tag element into ComboBox list
                      DisplaylistItem(Itemi, 1) = strTagPath
                      DisplaylistItem(Itemi, 2) = strTag
                      DisplaylistItem(Itemi, 3) = pagestr 'Put the tagname:screenname into ComboBox list for selecting
                      Itemi = Itemi + 1
                      
                      NumFile = NumFile + 1
                      
                      Tagnames(searchFile) = strTag  'The two arrays is to delete the repeated items from ComboBox list
                      ScreenName(searchFile) = pagestr ' The two arrays are similar with the caches of ComboBox list
                      searchFile = searchFile + 1 ' searchFile is pointer that alway point at the last item of the two arrays
                Else
                
                   If InStr(1, Tagnames(searchj), strTag) <> 0 And InStr(1, ScreenName(searchj), pagestr) <> 0 Then
                   
                      Exit For ' As The new tag element is the same with one item of two arrays, ignore it
                      
                   ElseIf searchj = searchFile Then ' As the two arrays do not contain the new finded tag element, add it into ComboBox
                      DisplaylistItem(Itemi, 1) = strTagPath
                      DisplaylistItem(Itemi, 2) = strTag
                      DisplaylistItem(Itemi, 3) = pagestr 'Put the tagname:screenname into ComboBox list for selecting
                      Itemi = Itemi + 1
                      
                      NumFile = NumFile + 1
                      
                      Tagnames(searchFile) = strTag
                      ScreenName(searchFile) = pagestr
                      searchFile = searchFile + 1
                                      
                   End If
                   

                End If
                
             Next

        End If
            
     
     Next
     
'********************************************************************************************************************
'For this part, you can delete duplicated items in string array, the array is a 2-D, like that: the first column is the path of search tags, the second column is the searched tags ,
'the other columns are about screens that contain the searched tags
'********************************************************************************************************************
     For i = 1 To NumFile

       strm = DisplaylistItem(i, 2)
       
   
       For j = i + 1 To NumFile
   
         If DisplaylistItem(j, 2) = strm Then
      
           DisplaylistItem(i, numItem) = DisplaylistItem(j, 3)
           
           numItem = numItem + 1
         
           For n = j To NumFile

              DisplaylistItem(n, 1) = DisplaylistItem(n + 1, 1)
              DisplaylistItem(n, 2) = DisplaylistItem(n + 1, 2)
              DisplaylistItem(n, 3) = DisplaylistItem(n + 1, 3)
            
           Next
         
           NumFile = NumFile - 1
           j = j - 1
      
        End If
   
      Next
   
      numItem = 4

    Next

    Call ArrToListBox  'Put all items in the first column of 2-D string array into SearchTag listbox

'/////////////////////////////////////////////////////////////
     
     
    Set oMyXmlNode = Nothing
    Set oMyXmlDoc = Nothing ' Release the cache of XML document
         
    Exit Sub
     
errHandler:

    Set oMyXmlNode = Nothing  ' Release the cache of XML document
    Set oMyXmlDoc = Nothing   ' Release the cache of XML document
    MsgBox ("Erro...Please check the path of search file !")
     
End Sub

Private Sub ArrToListBox()

   Dim indexArr As Integer

   TagListBox.Clear 'clear the listbox

   indexArr = 1

   Do While DisplaylistItem(indexArr, 2) <> ""

     TagListBox.AddItem DisplaylistItem(indexArr, 2) 'put searched tag into listbox
     indexArr = indexArr + 1

   Loop

End Sub


Private Sub TagListBox_Click()

   Dim TagListIndex As Integer
   
   TagIndexInDislayArray = 0
   TagListIndex = TagListBox.ListIndex + 1
   TagIndexInDislayArray = TagListIndex
   
   ScreenListBox.Clear

   i = 3  'point the third column of string array, which is the screens name

   Do While DisplaylistItem(TagListIndex, i) <> ""
   
     ScreenListBox.AddItem DisplaylistItem(TagListIndex, i) 'put all screens into screen listbox, which is based on searched tag that you select
     i = i + 1
   
   Loop

End Sub


'*****The purpose of this part is to command search function after you type tag/device name and press ENTER key at keyboard******
Private Sub TagInput_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If KeyCode = 13 Then
    
       Call Searchbutton_Click
    
    End If

End Sub


Private Sub InitialTagFaceplateArray()
'******* xxxTagFaceplate is 2-D array, the first column is tag mark which can distinguish tag, the second column is faceplate name which has relation with tagmark*********
'******* if you have more folds in HMI tag server, you can create xxxTagFaceplate array, then finish relation in this function**************
'******* ESD/FGS has another tag mark, you can add them into this function************************
ESDTagFaceplate(1, 1) = "AL"
ESDTagFaceplate(2, 1) = "BDY"
ESDTagFaceplate(3, 1) = "ESD"
ESDTagFaceplate(4, 1) = "LDY"
ESDTagFaceplate(5, 1) = "UY"
ESDTagFaceplate(6, 1) = "XS"
ESDTagFaceplate(7, 1) = "HPU"
ESDTagFaceplate(8, 1) = "HS"
ESDTagFaceplate(9, 1) = "PB"
ESDTagFaceplate(10, 1) = "TRIP01"
ESDTagFaceplate(11, 1) = "TRIP02"
ESDTagFaceplate(12, 1) = "TRIP03"
ESDTagFaceplate(13, 1) = "TSHH"
ESDTagFaceplate(14, 1) = "UA"
ESDTagFaceplate(15, 1) = "LIT"
ESDTagFaceplate(16, 1) = "LT"
ESDTagFaceplate(17, 1) = "PIT"
ESDTagFaceplate(18, 1) = "PT"
ESDTagFaceplate(19, 1) = "TIT"
ESDTagFaceplate(20, 1) = "SDY"
ESDTagFaceplate(1, 2) = "(og)esd_do_faceplate"
ESDTagFaceplate(2, 2) = "(og)esd_do_faceplate"
ESDTagFaceplate(3, 2) = "(og)esd_do_faceplate"
ESDTagFaceplate(4, 2) = "(og)esd_do_faceplate"
ESDTagFaceplate(5, 2) = "(og)esd_do_faceplate"
ESDTagFaceplate(6, 2) = "(og)esd_do_faceplate"
ESDTagFaceplate(7, 2) = "(og)esd_di_faceplate"
ESDTagFaceplate(8, 2) = "(og)esd_di_faceplate"
ESDTagFaceplate(9, 2) = "(og)esd_di_faceplate"
ESDTagFaceplate(10, 2) = "(og)esd_di_faceplate"
ESDTagFaceplate(11, 2) = "(og)esd_di_faceplate"
ESDTagFaceplate(12, 2) = "(og)esd_di_faceplate"
ESDTagFaceplate(13, 2) = "(og)esd_di_faceplate"
ESDTagFaceplate(14, 2) = "(og)esd_di_faceplate"
ESDTagFaceplate(15, 2) = "(og)esd_ai_faceplate"
ESDTagFaceplate(16, 2) = "(og)esd_ai_faceplate"
ESDTagFaceplate(17, 2) = "(og)esd_ai_faceplate"
ESDTagFaceplate(18, 2) = "(og)esd_ai_faceplate"
ESDTagFaceplate(19, 2) = "(og)esd_ai_faceplate"
ESDTagFaceplate(20, 2) = "(og)esd_do_faceplate"


FGSTagFaceplate(1, 1) = "FD"
FGSTagFaceplate(2, 1) = "GD"
FGSTagFaceplate(3, 1) = "GDX"
FGSTagFaceplate(4, 1) = "GDTR"
FGSTagFaceplate(5, 1) = "H2D"
FGSTagFaceplate(6, 1) = "PIT"
FGSTagFaceplate(7, 1) = "PT"
FGSTagFaceplate(8, 1) = "SD"
FGSTagFaceplate(9, 1) = "TIT"
FGSTagFaceplate(10, 1) = "GDR"
FGSTagFaceplate(11, 1) = "C2H2"
FGSTagFaceplate(12, 1) = "AL"
FGSTagFaceplate(13, 1) = "SDY"
FGSTagFaceplate(14, 1) = "STB"
FGSTagFaceplate(15, 1) = "SV"
FGSTagFaceplate(16, 1) = "US"
FGSTagFaceplate(17, 1) = "UY"
FGSTagFaceplate(18, 1) = "HD"
FGSTagFaceplate(19, 1) = "HS"
FGSTagFaceplate(20, 1) = "MFS"
FGSTagFaceplate(21, 1) = "MI"
FGSTagFaceplate(22, 1) = "MR"
FGSTagFaceplate(23, 1) = "PB"
FGSTagFaceplate(24, 1) = "PSH"
FGSTagFaceplate(25, 1) = "PSL"
FGSTagFaceplate(26, 1) = "UA"
FGSTagFaceplate(27, 1) = "XL"
FGSTagFaceplate(28, 1) = "ZSC"
FGSTagFaceplate(29, 1) = "ZSO"
FGSTagFaceplate(1, 2) = "(og)fgs_ai_faceplate"
FGSTagFaceplate(2, 2) = "(og)fgs_ai_faceplate"
FGSTagFaceplate(3, 2) = "(og)fgs_ai_faceplate"
FGSTagFaceplate(4, 2) = "(og)fgs_ai_faceplate"
FGSTagFaceplate(5, 2) = "(og)fgs_ai_faceplate"
FGSTagFaceplate(6, 2) = "(og)fgs_ai_faceplate"
FGSTagFaceplate(7, 2) = "(og)fgs_ai_faceplate"
FGSTagFaceplate(8, 2) = "(og)fgs_ai_faceplate"
FGSTagFaceplate(9, 2) = "(og)fgs_ai_faceplate"
FGSTagFaceplate(10, 2) = "(og)fgs_ai_faceplate"
FGSTagFaceplate(11, 2) = "(og)fgs_ai_faceplate"
FGSTagFaceplate(12, 2) = "(og)fgs_do_faceplate"
FGSTagFaceplate(13, 2) = "(og)fgs_do_faceplate"
FGSTagFaceplate(14, 2) = "(og)fgs_do_faceplate"
FGSTagFaceplate(15, 2) = "(og)fgs_do_faceplate"
FGSTagFaceplate(16, 2) = "(og)fgs_do_faceplate"
FGSTagFaceplate(17, 2) = "(og)fgs_do_faceplate"
FGSTagFaceplate(18, 2) = "(og)fgs_di_faceplate"
FGSTagFaceplate(19, 2) = "(og)fgs_di_faceplate"
FGSTagFaceplate(20, 2) = "(og)fgs_di_faceplate"
FGSTagFaceplate(21, 2) = "(og)fgs_di_faceplate"
FGSTagFaceplate(22, 2) = "(og)fgs_di_faceplate"
FGSTagFaceplate(23, 2) = "(og)fgs_di_faceplate"
FGSTagFaceplate(24, 2) = "(og)fgs_di_faceplate"
FGSTagFaceplate(25, 2) = "(og)fgs_di_faceplate"
FGSTagFaceplate(26, 2) = "(og)fgs_di_faceplate"
FGSTagFaceplate(27, 2) = "(og)fgs_di_faceplate"
FGSTagFaceplate(28, 2) = "(og)fgs_di_faceplate"
FGSTagFaceplate(29, 2) = "(og)fgs_di_faceplate"



End Sub

