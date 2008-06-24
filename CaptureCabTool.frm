VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Capture Cab Tool"
   ClientHeight    =   5880
   ClientLeft      =   1695
   ClientTop       =   2355
   ClientWidth     =   8160
   Icon            =   "CaptureCabTool.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   5880
   ScaleWidth      =   8160
   Begin VB.CheckBox chkXtrata 
      Caption         =   "Xtrata"
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   1080
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CheckBox chkSamples 
      Caption         =   "Prepare Samples for Testing"
      Height          =   255
      Left            =   5400
      TabIndex        =   10
      Top             =   480
      Value           =   1  'Checked
      Width           =   2655
   End
   Begin VB.CheckBox chkTemp 
      Caption         =   "{System TMP/TEMP}"
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   720
      Width           =   4815
   End
   Begin VB.CheckBox chkImagePath 
      Caption         =   "{Server UNC Path}"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   480
      Value           =   1  'Checked
      Width           =   4815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   8280
      TabIndex        =   6
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   8280
      TabIndex        =   5
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   8280
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdRepack 
      Caption         =   "Repack"
      Height          =   495
      Left            =   8520
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Process"
      Height          =   495
      Left            =   8520
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   495
      Left            =   8520
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   4455
      Left            =   120
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "CaptureCabTool.frx":000C
      Top             =   1320
      Width           =   7935
   End
   Begin VB.Label lblRemove 
      Caption         =   "Remove:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Replace Temp Image Path with:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
 'unique values from array declarations
Const ERR_BAD_PARAMETER = "Array parameter required"
Const ERR_BAD_TYPE = "Invalid Type"
Const ERR_BP_NUMBER = 20000
Const ERR_BT_NUMBER = 20001

'show open through api
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias _
         "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias _
         "GetSaveFileNameA" (pSavefilename As OPENFILENAME) As Long

    Private Type OPENFILENAME
      lStructSize As Long
      hwndOwner As Long
      hInstance As Long
      lpstrFilter As String
      lpstrCustomFilter As String
      nMaxCustFilter As Long
      nFilterIndex As Long
      lpstrFile As String
      nMaxFile As Long
      lpstrFileTitle As String
      nMaxFileTitle As Long
      lpstrInitialDir As String
      lpstrTitle As String
      Flags As Long
      nFileOffset As Integer
      nFileExtension As Integer
      lpstrDefExt As String
      lCustData As Long
      lpfnHook As Long
      lpTemplateName As String
    End Type

Public Sub FullProcess(Filename As String)
On Error GoTo Err:
If LCase(Mid(Filename, Len(Filename) - 3, 4)) = ".cab" Then
    SourceFilename = Mid(Filename, InStrRev(Filename, "\") + 1, Len(Filename) - InStrRev(Filename, "\"))
    Log vbNewLine & vbNewLine & Source
    
    SamplesDir = XMLDir & "\" & SourceFilename '& "\" 'matters to have end slash
    Dim fso As New FileSystemObject
    'fso.CreateFolder App.Path & "\cabexport"
    If Not fso.FolderExists(XMLDir) Then fso.CreateFolder XMLDir
    If Not fso.FolderExists(XMLDir & "\XMLBackup") Then fso.CreateFolder XMLDir & "\XMLBackup"
    If fso.FolderExists(SamplesDir) Then fso.DeleteFolder SamplesDir
    fso.CreateFolder SamplesDir
    
    On Error Resume Next
    Dim fldXMLDir As Folder
    Dim fXML As File
    Set fldXMLDir = fso.GetFolder(XMLDir)
    'this is ugly but fso.movefile just errors instead of overwriting
    'any file in xml folder, delete in backup folder
    For Each fXML In fldXMLDir.Files
        fso.DeleteFile XMLDir & "\XMLBackup\" & fXML.Name
    Next
    'then copy from xml folder to backup
    fso.MoveFile XMLDir & "\*.xml", XMLDir & "\XMLBackup"
    On Error GoTo 0

    UnpackCab Filename
    Process
    RepackCab Filename & "-test.cab"
    If chkSamples.Value = 1 Then Samples
Else
    MsgBox "Not a .cab file! (" & Filename & ")"
End If
Err:
If Err.Number > 0 Then
    MsgBox Err.Description & " (" & Err.Source & ")", vbCritical, Err.Number
    End
End If

End Sub

Public Sub UnpackCab(Cab As String)
Dim UnpackCommand As String
' if it errors on deleting the folder don't worry about it
On Error Resume Next
Dim ltmp As Long
Dim fso As New FileSystemObject
fso.DeleteFolder TempPath & "\cabexport"
On Error GoTo 0
'kfxcabar does not create the dir
fso.CreateFolder TempPath & "\cabexport"

'command to use ms cab tools
'UnpackCommand = """" & "extract.exe"" /Y /E /L " & TempPath & "\cabexport """ & Cab & """"

'command to use kofax cab tool, works as long as capture is install with bin on PATH variable, or kfxcabar in app path
UnpackCommand = "KfxCabAr.exe E """ & Cab & """ """ & TempPath & "\cabexport"""

Log "Uncabbing with command:" & UnpackCommand
SyncShell UnpackCommand

CabLog vbNewLine & vbNewLine & Source

End Sub

Public Sub RepackCab(Cab As String)
Dim RepackCommand As String
Dim CabNoEx As String 'No extension only used for packing (not unpacking) with kfxcabar
CabNoEx = Mid(Cab, 1, Len(Cab) - 4)
' if it errors on moving the log don't worry about it
On Error Resume Next
'Dim fso As New FileSystemObject
'fso.MoveFile App.Path & "\cabtool.log", App.Path & "\cabexport\cabtool.log"

On Error GoTo 0
'command to use ms cab tools
' the -P strips cabexport from the path within the cab
RepackCommand = """" & "cabarc.exe"" -p -P " & TempPath & "\cabexport\ -r -i 130 N """ & Cab & """ """ & TempPath & "\cabexport\*.*"""
'command to use kofax cab tool, works as long as capture is install with bin on PATH variable, or kfxcabar in app path
RepackCommand = "KfxCabAr.exe B """ & CabNoEx & """ """ & TempPath & "\cabexport"""

Log "Recabbing with command:" & RepackCommand
SyncShell RepackCommand
End Sub

'If there is an error accessing the attribute, then it doesn't exist
Public Function AttribExists(node As IXMLDOMNode, AttribName As String) As Boolean
On Error GoTo Err:

AttribExists = False
Dim Temp As String
Temp = node.Attributes.getNamedItem(AttribName).nodeValue
AttribExists = True

Err:
If Err.Number > 0 Then
    AttribExists = False
    CabLog "Attribute " & AttribName & " not present."
End If
End Function

Public Sub Process()
Dim doc As DOMDocument30
Dim node As IXMLDOMNode
Dim nodes As IXMLDOMNodeList
Dim attrib As IXMLDOMNode
Dim PreLog As String


Set doc = New DOMDocument30
doc.async = False
Log "Loading: " & TempPath & "\cabexport\Admin.xml"
doc.Load TempPath & "\cabexport\Admin.xml"

Set node = doc.selectSingleNode("//AscentCaptureSetup")
On Error Resume Next
CabLog "Exported from version " & DatabaseVersions(node.Attributes.getNamedItem("DatabaseVersion").nodeValue)
CabLog "Database version " & node.Attributes.getNamedItem("DatabaseVersion").nodeValue
On Error GoTo 0

'Remove release script
PreLog = "Removing Release Scripts:" & vbNewLine
Set nodes = doc.selectNodes("//ReleaseScript")
For Each node In nodes
    If ItemExists(node.Attributes.getNamedItem("Name").nodeValue, AllowedReleaseScripts) = False Then
        node.parentNode.removeChild node
        CabLog PreLog & "  """ & _
        node.Attributes.getNamedItem("Name").nodeValue & """ from cab."
        PreLog = ""
    End If
Next


'Remove release script config
PreLog = "Removing Release Script Configuration:" & vbNewLine
Set nodes = doc.selectNodes("//ReleaseSetup")
For Each node In nodes
    If ItemExists(node.Attributes.getNamedItem("ReleaseScriptName").nodeValue, AllowedReleaseScripts) = False Then
        node.parentNode.removeChild node
        CabLog PreLog & "  """ _
        & node.Attributes.getNamedItem("ReleaseScriptName").nodeValue & _
        """ from Doc Class """ & node.Attributes.getNamedItem("DocumentClassName").nodeValue & """."
        PreLog = ""
    End If
Next

'Remove module process
PreLog = "Removing Custom Module Processes:" & vbNewLine
Set nodes = doc.selectNodes("//Process[@ModuleID]") 'any Process node with a ModuleID
For Each node In nodes
    If ItemExists(node.Attributes.getNamedItem("ModuleID").nodeValue, AllowedModules) = False Then
        CabLog PreLog & "  """ & _
        node.Attributes.getNamedItem("ModuleID").nodeValue & """ from cab."
        node.parentNode.removeChild node
        PreLog = ""
    End If
Next

'Remove module (queue)
PreLog = "Removing Custom Modules:" & vbNewLine
Set nodes = doc.selectNodes("//Module[@ModuleID]")
For Each node In nodes
    If node.Attributes.getNamedItem("ModuleID").nodeValue = "AC.XtrataSvr" Then CabLog "NOTE: Xtrata is needed for Batch Class: " _
    & node.parentNode.parentNode.Attributes.getNamedItem("Name").nodeValue
    If ItemExists(node.Attributes.getNamedItem("ModuleID").nodeValue, AllowedModules) = False Then
        CabLog PreLog & "  """ & _
        node.Attributes.getNamedItem("ModuleID").nodeValue & _
        """ from Batch Class """ & node.parentNode.parentNode.Attributes.getNamedItem("Name").nodeValue & """."
        node.parentNode.removeChild node
        PreLog = ""
    End If
Next

'Remove workflow agent
PreLog = "Removing Workflow Agents:" & vbNewLine
Set nodes = doc.selectNodes("//AssignedWorkflowAgent")
For Each node In nodes
    If ItemExists(node.Attributes.getNamedItem("ModuleID").nodeValue, AllowedWorkflowAgents) = False Then
        CabLog PreLog & "  """ _
        & node.Attributes.getNamedItem("ModuleID").nodeValue & _
        """ from Batch Class """ & node.parentNode.parentNode.Attributes.getNamedItem("Name").nodeValue & """."
        node.parentNode.removeChild node
        PreLog = ""
    End If
Next

PreLog = "Removing Custom Modules From Foldering, Start-Sort, and Partial Batch Release:" & vbNewLine _
& "(These settings may need to be verified on the Foldering and Advanced tabs of the Batch Class Properties.)" & vbNewLine
Set nodes = doc.selectNodes("//BatchClass")
For Each node In nodes
    If AttribExists(node, "FolderCreationStartModuleID") Then
         If ItemExists(node.Attributes.getNamedItem("FolderCreationStartModuleID").nodeValue, AllowedModules) = False And node.Attributes.getNamedItem("FolderCreationStartModuleID").nodeValue <> "0" Then
             CabLog PreLog & "  WARNING! Auto-Foldering was set to """ & _
             node.Attributes.getNamedItem("FolderCreationStartModuleID").nodeValue & """ and has now been set to (none)."
             Set attrib = doc.createAttribute("FolderCreationStartModuleID")
             attrib.Text = "0"
             node.Attributes.setNamedItem attrib
             PreLog = ""
         End If
    End If
    If AttribExists(node, "SortStartModuleID") Then
        If ItemExists(node.Attributes.getNamedItem("SortStartModuleID").nodeValue, AllowedModules) = False And node.Attributes.getNamedItem("SortStartModuleID").nodeValue <> "0" Then
            CabLog PreLog & "  WARNING! Sort Options was set to """ & _
            node.Attributes.getNamedItem("SortStartModuleID").nodeValue & """ and has now been set to (none)."
            Set attrib = doc.createAttribute("SortStartModuleID")
            attrib.Text = "0"
            node.Attributes.setNamedItem attrib
            PreLog = ""
        End If
    End If
    If AttribExists(node, "AdvanceBatchesInErrorModuleID") Then
        If ItemExists(node.Attributes.getNamedItem("AdvanceBatchesInErrorModuleID").nodeValue, AllowedModules) = False And node.Attributes.getNamedItem("AdvanceBatchesInErrorModuleID").nodeValue <> "0" Then
            CabLog PreLog & "  WARNING! Partial Batch Release was set to """ & _
            node.Attributes.getNamedItem("AdvanceBatchesInErrorModuleID").nodeValue & """ and has now been set to (none)."
            Set attrib = doc.createAttribute("AdvanceBatchesInErrorModuleID")
            attrib.Text = "0"
            node.Attributes.setNamedItem attrib
            PreLog = ""
        End If
    End If
Next

Dim NewPath As String
'replace image path
If chkImagePath.Value = 1 Or chkTemp.Value = 1 Then
    If chkTemp.Value = 1 Then NewPath = TempPath
    If chkImagePath.Value = 1 Then NewPath = SVPath & "\Images"
    
    PreLog = "Replacing Image Path:" & vbNewLine
    Set nodes = doc.selectNodes("//BatchClass")
    For Each node In nodes
        CabLog PreLog & "  """ _
        & node.Attributes.getNamedItem("ImageDirectory").nodeValue & _
        """ from Batch Class """ & node.Attributes.getNamedItem("Name").nodeValue & _
        """ with """ & NewPath & """."
        PreLog = ""
        Set attrib = doc.createAttribute("ImageDirectory")
        attrib.Text = NewPath
        node.Attributes.setNamedItem attrib
    Next
End If

doc.save TempPath & "\cabexport\Admin.xml"
End Sub
       
Public Sub XMLSettings()
Dim doc As DOMDocument30
Dim node As IXMLDOMNode
Dim nodes As IXMLDOMNodeList
Dim attrib As IXMLDOMNode

On Error Resume Next 'If xml is malformed or duplicates are being added, just keep going

Set doc = New DOMDocument30
doc.async = False
Form1.Text1.Text = Form1.Text1.Text & vbNewLine & "Loaded settings from CaptureCabTool.xml" & vbNewLine
doc.Load App.Path & "\CaptureCabTool.xml"


Set nodes = doc.selectNodes("//DatabaseVersion")
For Each node In nodes
    DatabaseVersions.Add node.Attributes.getNamedItem("CaptureVersion").nodeValue, node.Attributes.getNamedItem("DatabaseVersion").nodeValue
    'MsgBox node.Attributes.getNamedItem("CaptureVersion").nodeValue & "," & node.Attributes.getNamedItem("DatabaseVersion").nodeValue
Next

Set nodes = doc.selectNodes("//ReleaseScript")
For Each node In nodes
    AllowedReleaseScripts.Add node.Attributes.getNamedItem("Name").nodeValue, node.Attributes.getNamedItem("Name").nodeValue
    'MsgBox node.Attributes.getNamedItem("Name").nodeValue & "," & node.Attributes.getNamedItem("Name").nodeValue
Next

Set nodes = doc.selectNodes("//Module")
For Each node In nodes
    AllowedModules.Add node.Attributes.getNamedItem("Name").nodeValue, node.Attributes.getNamedItem("Name").nodeValue
    'MsgBox node.Attributes.getNamedItem("Name").nodeValue & "," & node.Attributes.getNamedItem("Name").nodeValue
Next

Err.Number = 0

End Sub

Public Sub SaveXMLSettings()
Dim doc As DOMDocument30
Dim root As IXMLDOMNode
Dim node As IXMLDOMNode
Dim nodes As IXMLDOMNodeList
Dim attrib As IXMLDOMNode


Set doc = New DOMDocument30
doc.async = False

Set root = doc.createElement("CaptureCabTool")
Set root = doc.appendChild(root)

'for each

doc.save App.Path & "\CaptureCabToolExport.xml"

End Sub

Private Sub chkImagePath_Click()
If chkImagePath.Value = 1 Then
    chkTemp.Value = 0
End If
End Sub

Private Sub chkTemp_Click()
If chkTemp.Value = 1 Then
    chkImagePath.Value = 0
End If

End Sub

Private Sub chkXtrata_Click()
On Error GoTo Err:
If chkXtrata.Value = 1 Then
    AllowedModules.Remove "AC.XtrataSvr"
Else
    AllowedModules.Add "AC.XtrataSvr", "AC.XtrataSvr"
End If
Err:
If Err.Number > 0 Then
    MsgBox Err.Description, vbCritical, Err.Source
End If
End Sub

Private Sub cmdOpen_Click()
Dim OpenFile As OPENFILENAME
Dim lReturn As Long
Dim sFilter As String
OpenFile.lStructSize = Len(OpenFile)
OpenFile.hwndOwner = Form1.hWnd
OpenFile.hInstance = App.hInstance
sFilter = "Cab Files (*.cab)" & Chr(0) & "*.CAB" & Chr(0)
OpenFile.lpstrFilter = sFilter
OpenFile.nFilterIndex = 1
OpenFile.lpstrFile = String(257, 0)
OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
OpenFile.lpstrFileTitle = OpenFile.lpstrFile
OpenFile.nMaxFileTitle = OpenFile.nMaxFile
OpenFile.lpstrInitialDir = App.Path
OpenFile.lpstrTitle = "Open an exported cab file"
OpenFile.Flags = 0
lReturn = GetOpenFileName(OpenFile)
If lReturn = 0 Then
   
Else
   UnpackCab (OpenFile.lpstrFile)
End If
End Sub

Private Sub Command1_Click()
GetCabInfo
End Sub

Private Sub cmdRepack_Click()
Dim Cab As String
Dim OpenFile As OPENFILENAME
Dim lReturn As Long
Dim sFilter As String
OpenFile.lStructSize = Len(OpenFile)
OpenFile.hwndOwner = Form1.hWnd
OpenFile.hInstance = App.hInstance
sFilter = "Cab Files (*.cab)" & Chr(0) & "*.CAB" & Chr(0)
OpenFile.lpstrFilter = sFilter
OpenFile.nFilterIndex = 1
OpenFile.lpstrFile = String(257, 0)
OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
OpenFile.lpstrFileTitle = OpenFile.lpstrFile
OpenFile.nMaxFileTitle = OpenFile.nMaxFile
OpenFile.lpstrInitialDir = App.Path
OpenFile.lpstrTitle = "Save repacked cab file"
OpenFile.Flags = 0
OpenFile.lpstrDefExt = ".cab"
lReturn = GetSaveFileName(OpenFile)
If lReturn = 0 Then
   
Else
    'to deal with nulls from fixed length string
    Cab = Mid(OpenFile.lpstrFile, 1, InStr(1, OpenFile.lpstrFile, Chr(0)) - 1)
    RepackCab Cab
End If
End Sub

Private Sub Command2_Click()
Process
End Sub

Private Sub Command3_Click()
Samples
End Sub


Private Sub Command4_Click()
MsgBox GetRegString(HKEY_LOCAL_MACHINE, "Software\Kofax Image Products\Ascent Capture\3.0", "ServerPath")

End Sub

Private Sub Form_Load()
'On Error GoTo Err
Text1.Text = "Capture Cab Tool " & App.Major & "." & App.Minor & "." & App.Revision & " - Stephen Klancher" & vbNewLine & Text1.Text

XMLDir = "C:\ACXMLAID"

TempPath = IIf(Environ$("tmp") <> "", Environ$("tmp"), Environ$("temp"))
If TempPath = "" Then
    MsgBox "Please ensure that either TMP or TEMP is set as an environment variable as this is used as the application working directory.", vbCritical
    End
End If
SVPath = GetRegString(HKEY_LOCAL_MACHINE, "Software\Kofax Image Products\Ascent Capture\3.0", "ServerPath")

If SVPath = "" Then
    chkImagePath.Value = 0
    chkImagePath.Enabled = False
    Text1.Text = Text1.Text & vbNewLine & "======  NOTE  ======" & vbNewLine _
    & "Capture was not detected as installed on this system.  To use this application ensure that you have copied " _
    & "KfxCabAr.exe from the bin directory of a Capture install to either the same directory as this app, or to " _
    & "somewhere on the path variable." & vbNewLine & "====== ==== ======"
End If

chkImagePath.Caption = SVPath & "\Images"
chkTemp.Caption = TempPath

DatabaseVersions.Add "8.0", "25"
DatabaseVersions.Add "7.5", "23"
DatabaseVersions.Add "7.0", "22"
DatabaseVersions.Add "6.1", "17"
DatabaseVersions.Add "6.0?", "8"


AllowedReleaseScripts.Add "Ascent Capture Database", "Ascent Capture Database"
AllowedReleaseScripts.Add "Ascent Capture Text", "Ascent Capture Text"

'Allow all normal modules
AllowedModules.Add "scan.exe", "scan.exe"
AllowedModules.Add "fp.exe", "fp.exe"
AllowedModules.Add "index.exe", "index.exe"
AllowedModules.Add "verify.exe", "verify.exe"
AllowedModules.Add "kfxpdf.exe", "kfxpdf.exe"
AllowedModules.Add "ocr.exe", "ocr.exe"
AllowedModules.Add "qc.exe", "qc.exe"
AllowedModules.Add "release.exe", "release.exe"
'Allow Xtrata
'AllowedModules.Add "AC.XtrataSvr", "AC.XtrataSvr"

Dim fso As FileSystemObject
Set fso = New FileSystemObject
If fso.FileExists(App.Path & "\CaptureCabTool.xml") Then
    XMLSettings
End If

'process commandline files one by one
Dim i As Integer
For i = 1 To GetParamCount
    Source = GetParam(i)
    'folder of source cab with ending \
    SourceDir = Mid(Source, 1, InStrRev(Source, "\"))
    FullProcess Source
Next i


Err:
If Err.Number > 0 Then MsgBox Err.Description, vbCritical, Err.Number
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
For Each File In Data.Files
    Source = CStr(File)
    'folder of source cab with ending \
    SourceDir = Mid(Source, 1, InStrRev(Source, "\"))
    FullProcess Source
Next
End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub
