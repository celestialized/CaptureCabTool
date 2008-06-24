Attribute VB_Name = "Module1"
Public Declare Function WaitForCmd Lib "msvcrt.dll" (ByVal sCommand As String) As Long
Declare Function WaitForSingleObject Lib "Kernel32" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
Declare Function InputIdle Lib "user32" Alias "WaitForInputIdle" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
Declare Function CreateProcessA Lib "Kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long

Public Type STARTUPINFO
    cb              As Long
    lpReserved      As Long
    lpDesktop       As Long
    lpTitle         As Long
    dwX             As Long
    dwY             As Long
    dwXSize         As Long
    dwYSize         As Long
    dwXCountChars   As Long
    dwYCountChars   As Long
    dwFillAttribute As Long
    dwFlags         As Long
    wShowWindow     As Integer
    cbReserved2     As Integer
    lpReserved2     As Long
    hStdInput       As Long
    hStdOutput      As Long
    hStdError       As Long
End Type

Public Type PROCESS_INFORMATION
    hProcess    As Long
    hThread     As Long
    dwProcessID As Long
    dwThreadID  As Long
End Type

Public OpenCab As CabFile

Public Type CabFile
    File As String
    DatabaseVersion As Integer
    BatchClasses As Collection
    DocumentClasses As Collection
End Type

Public Type BatchClass
    Name As String
    DocumentClasses As Collection
End Type


Public Type FormTypeArray
    Name As String
    Samples() As String
End Type

Public Type DocClassArray
    Name As String
    FormTypes() As FormTypeArray
End Type

Public Type BatchClassArray
    Name As String
    DocumentClasses() As DocClassArray
End Type



Public Type DocumentClass
    Name As String
    FormTypes As Collection
End Type

Public Type FormType
    Name As String
    SamplePages As Collection
End Type

Public DocClasses() As DocClassArray

Public AllowedModules As New Collection
Public AllowedReleaseScripts As New Collection
Public AllowedWorkflowAgents As New Collection
Public DatabaseVersions As New Collection

Public TempPath As String
Public SVPath As String
Public SourceDir As String
Public Source As String
Public SourceFilename As String
Public XMLDir As String
Public SamplesDir As String


'Shell Constants
Global Const NORMAL_PRIORITY_CLASS      As Long = &H20&
Global Const INFINITE                   As Long = -1&

Global Const STATUS_WAIT_0              As Long = &H0
Global Const STATUS_ABANDONED_WAIT_0    As Long = &H80
Global Const STATUS_USER_APC            As Long = &HC0
Global Const STATUS_TIMEOUT             As Long = &H102
Global Const STATUS_PENDING             As Long = &H103

Global Const WAIT_FAILED                As Long = &HFFFFFFFF
Global Const WAIT_OBJECT_0              As Long = STATUS_WAIT_0
Global Const WAIT_TIMEOUT               As Long = STATUS_TIMEOUT

Global Const WAIT_ABANDONED             As Long = STATUS_ABANDONED_WAIT_0
Global Const WAIT_ABANDONED_0           As Long = STATUS_ABANDONED_WAIT_0

Global Const WAIT_IO_COMPLETION         As Long = STATUS_USER_APC
Global Const STILL_ACTIVE               As Long = STATUS_PENDING

Public Function StripInvalidChar(Filename As String)
StripInvalidChar = Replace(Filename, "\", "")
StripInvalidChar = Replace(StripInvalidChar, "/", "")
StripInvalidChar = Replace(StripInvalidChar, ":", "")
StripInvalidChar = Replace(StripInvalidChar, "*", "")
StripInvalidChar = Replace(StripInvalidChar, """", "")
StripInvalidChar = Replace(StripInvalidChar, "<", "")
StripInvalidChar = Replace(StripInvalidChar, ">", "")
StripInvalidChar = Replace(StripInvalidChar, "?", "")
StripInvalidChar = Replace(StripInvalidChar, "|", "")

End Function


''==============================================================================
''Code flow routines:

Public Function SyncShell(CommandLine As String, Optional Timeout As Long, _
    Optional WaitForInputIdle As Boolean, Optional Hide As Boolean = False) As Boolean

    Dim hProcess As Long

    Const STARTF_USESHOWWINDOW As Long = &H1
    Const SW_HIDE As Long = 0
    
    Dim Ret As Long
    Dim nMilliseconds As Long

    If Timeout > 0 Then
        nMilliseconds = Timeout
    Else
        nMilliseconds = INFINITE
    End If

    hProcess = StartProcess(CommandLine, Hide)

    If WaitForInputIdle Then
        'Wait for the shelled application to finish setting up its UI:
        Ret = InputIdle(hProcess, nMilliseconds)
    Else
        'Wait for the shelled application to terminate:
        Ret = WaitForSingleObject(hProcess, nMilliseconds)
    End If

    CloseHandle hProcess

    'Return True if the application finished. Otherwise it timed out or erred.
    SyncShell = (Ret = WAIT_OBJECT_0)
End Function


Public Function StartProcess(CommandLine As String, Optional Hide As Boolean = False) As Long
    Const STARTF_USESHOWWINDOW As Long = &H1
    Const SW_HIDE As Long = 0
    
    Dim proc As PROCESS_INFORMATION
    Dim Start As STARTUPINFO

    'Initialize the STARTUPINFO structure:
    Start.cb = Len(Start)
    If Hide Then
        Start.dwFlags = STARTF_USESHOWWINDOW
        Start.wShowWindow = SW_HIDE
    End If
    'Start the shelled application:
    CreateProcessA 0&, CommandLine, 0&, 0&, 1&, _
        NORMAL_PRIORITY_CLASS, 0&, 0&, Start, proc

    StartProcess = proc.hProcess
End Function


Public Sub Log(msg As String)
Dim FileNum As Integer
FileNum = FreeFile
Open App.Path & "\cabtool.log" For Append As #FileNum
Print #FileNum, Format(Now, "yyyy-mm-dd Hh:Nn.Ss - ") & msg
Close #FileNum
'Form1.Text1.Text = Form1.Text1.Text & vbNewLine & msg
End Sub

Public Sub CabLog(msg As String)
Dim FileNum As Integer
FileNum = FreeFile
Open TempPath & "\cabexport\cabtool.log" For Append As #FileNum
Print #FileNum, Format(Now, "yyyy-mm-dd Hh:Nn.Ss - ") & msg
Close #FileNum
Form1.Text1.Text = Form1.Text1.Text & vbNewLine & msg
End Sub



Public Function ItemExists(Item As String, Col As Collection)
On Error GoTo NotPresent
Dim Value As Variant

Value = Col(Item)
ItemExists = True

NotPresent:
If Err.Number > 0 Then ItemExists = False
End Function

Public Sub ClearCollection(ByRef Col As Collection)
For Each Item In Col
    Col.Remove Item
Next
End Sub


Public Sub GetCabInfo()
Dim doc As DOMDocument30
Dim node As IXMLDOMNode
Dim subnode As IXMLDOMNode
Dim nodes As IXMLDOMNodeList
Dim subnodes As IXMLDOMNodeList
Dim NewBatchClass As BatchClass
Dim NewDocClass As DocumentClass

'Set NewBatchClass = New BatchClass
'Set NewDocClass = New NewDocClass



NewDocClass.FormTypes.Add "test"
ClearCollection NewDocClass.FormTypes


Set doc = New DOMDocument30
doc.async = False
doc.Load App.Path & "\cabexport\Admin.xml"

OpenCab.DatabaseVersion = doc.selectSingleNode("/AscentCaptureSetup").Attributes.getNamedItem("DatabaseVersion").nodeValue
If ItemExists(CStr(OpenCab.DatabaseVersion), DatabaseVersions) Then
    Log "Cab from Capture " & DatabaseVersions(CStr(OpenCab.DatabaseVersion)) & " (DatabaseVersion " & OpenCab.DatabaseVersion & ")"
Else
    Log "Cab from unknown version of Capture (DatabaseVersion " & OpenCab.DatabaseVersion & ")"
End If

Set nodes = doc.selectNodes("//BatchClass")
For Each node In nodes
    NewBatchClass.Name = node.Attributes.getNamedItem("Name").nodeValue
    'This depends on DocumentClassLinks being the third child, would be better to get to it a different way
    Set subnodes = node.childNodes(2).childNodes
    For Each subnode In subnodes
        MsgBox subnode.Attributes.getNamedItem("DocumentClassName").nodeValue
        
    Next
Next

End Sub

Public Function GetSamples(DCName As String) As String()
'function no longer needed
Dim i As Integer
Dim temp() As String
ReDim temp(0)
GetSamples = temp
For i = 0 To UBound(DocClasses)
    If DocClasses(i).Name = DCName Then
        'GetSamples = DocClasses(i).Samples
        Exit For
    End If
Next i
End Function

Public Function GetDocNum(DCName As String) As Integer
Dim i As Integer
For i = 0 To UBound(DocClasses)
    If DocClasses(i).Name = DCName Then
        GetDocNum = i
        Exit For
    End If
Next i
End Function

'This whole thing is messy and should be redone
'...but won't be.
Public Sub Samples()

Dim doc As DOMDocument30
Dim HereXML As DOMDocument30
Dim ThereXML As DOMDocument30
Dim TImport As IXMLDOMNode
Dim TBatches As IXMLDOMNode
Dim TBatch As IXMLDOMNode

Dim BatchName As IXMLDOMNode
Dim BatchClass As IXMLDOMNode
Dim Priority As IXMLDOMNode
Dim Processed As IXMLDOMNode

Dim Documents As IXMLDOMNode
Dim Document As IXMLDOMNode

Dim FormType As IXMLDOMNode

Dim Pages As IXMLDOMNode
Dim Page As IXMLDOMNode

Dim ImportFileName As IXMLDOMNode

Dim Here As IXMLDOMNode
Dim TempNode As IXMLDOMNode
Dim node As IXMLDOMNode
Dim subnode As IXMLDOMNode
Dim nodes As IXMLDOMNodeList
Dim subnodes As IXMLDOMNodeList
Dim BCArray() As String
Dim FTNum As Integer
Dim SNum As Integer
Dim DCNum As Integer
Dim NumSamples As Integer
Dim SampleDir As Integer

Dim SamplesArray() As String

Dim BCName As String
Dim BCFileName As String

Dim f As Integer
Dim s As Integer

Dim i As Integer
Dim fso As New FileSystemObject
    

DCNum = 0



Set doc = New DOMDocument30
doc.async = False
Log "Loading: " & TempPath & "\cabexport\Admin.xml for XML AI creation."
doc.Load TempPath & "\cabexport\Admin.xml"


Set nodes = doc.selectNodes("//DocumentClass")
'each doc class
For Each node In nodes
    ReDim Preserve DocClasses(DCNum)
    DocClasses(DCNum).Name = node.Attributes.getNamedItem("Name").nodeValue

    
    FTNum = 0
    'not sure if there is a way to select decendents from a specific node without another search
    'but this should pick formtypes from the current doc class
    Set subnodes = doc.selectNodes("//DocumentClass[@Name='" & DocClasses(DCNum).Name & "']/FormTypes/FormType")
    'each formtype
    For Each subnode In subnodes
        SNum = 0
        
        NumSamples = subnode.Attributes.getNamedItem("SamplePageCount").nodeValue
        SampleDir = subnode.Attributes.getNamedItem("SampleImageDirectoryNumber").nodeValue
        ReDim Preserve DocClasses(DCNum).FormTypes(FTNum)
        DocClasses(DCNum).FormTypes(FTNum).Name = subnode.Attributes.getNamedItem("Name").nodeValue
        'each sample page
        ReDim Preserve DocClasses(DCNum).FormTypes(FTNum).Samples(SNum) 'so we don't leave samples empty, even if blank
        For i = 1 To NumSamples
            ReDim Preserve DocClasses(DCNum).FormTypes(FTNum).Samples(SNum)
            'leading zeros format does not work on hex because it is not numeric
            DocClasses(DCNum).FormTypes(FTNum).Samples(SNum) = Replace(Format(Hex(SampleDir), "@@@@@@@@"), " ", "0") & "\" & i & ".tif"
            SNum = SNum + 1
        Next i
        
        FTNum = FTNum + 1
    Next
    
    DCNum = DCNum + 1
Next

'Dim HereFile As Integer
Dim ThereFile As Integer
Dim SampleFileNumber As Integer

'create folders, ignore error if they exist
'== moved to form1.fullprocess

Set nodes = doc.selectNodes("//BatchClass")
'for each batch class
For Each node In nodes
    BCName = node.Attributes.getNamedItem("Name").nodeValue
    BCFileName = StripInvalidChar(BCName)
    
    
    'xml every batch class
    Set ThereXML = New DOMDocument30
    ThereXML.async = False
    ThereXML.preserveWhiteSpace = True

    Set TImport = ThereXML.createElement("ImportSession")
    Set TImport = ThereXML.appendChild(TImport)
    Set TBatches = ThereXML.createElement("Batches")
    Set TBatches = TImport.appendChild(TBatches)

    Set TBatch = ThereXML.createElement("Batch")
    Set TBatch = TBatches.appendChild(TBatch)
    
    Set BatchName = ThereXML.createAttribute("Name")
    TBatch.Attributes.setNamedItem BatchName
    BatchName.Text = ""
    
    Set BatchClass = ThereXML.createAttribute("BatchClassName")
    TBatch.Attributes.setNamedItem BatchClass
    BatchClass.Text = BCName
    
    Set Priority = ThereXML.createAttribute("Priority")
    TBatch.Attributes.setNamedItem Priority
    Priority.Text = "1"
    
    Set Processed = ThereXML.createAttribute("Processed")
    TBatch.Attributes.setNamedItem Processed
    Processed.Text = "0"
    
    Set Documents = ThereXML.createElement("Documents")
    Set Documents = TBatch.appendChild(Documents)
    
    

    SampleFileNumber = 0
    
    'HereFile = FreeFile
    'Open Source & "-Samples\" & BCName & "(From Here).ini" For Output As #HereFile
    '*ThereFile = FreeFile
    '*Open Source & "-Samples\" & BCFileName & ".ini" For Output As #ThereFile
    'Print #HereFile, "[Options]" & vbNewLine & "BatchClass=" & BCName & vbNewLine & "BatchName=" & vbNewLine & "Description="
    '*Print #ThereFile, "[Options]" & vbNewLine & "BatchClass=" & BCName & vbNewLine & "BatchName=" & vbNewLine & "Description="
    
    'documentclasslinks
    Set subnodes = doc.selectNodes("//BatchClass[@Name='" & BCName & "']/DocumentClassLinks/DocumentClassLink")
    'Set subnodes = node.childNodes(2).childNodes
    For Each subnode In subnodes
        'every doc class
        Dim DocNum As Integer
        DocNum = GetDocNum(subnode.Attributes.getNamedItem("DocumentClassName").nodeValue)
        'every form type in doc class
        For f = 0 To UBound(DocClasses(DocNum).FormTypes) 'need to iterate by name in xml, not sequentially
            'create a document for each formtype
            Set Document = ThereXML.createElement("Document")
            Set Document = Documents.appendChild(Document)
            
            Set FormType = ThereXML.createAttribute("FormTypeName")
            Document.Attributes.setNamedItem FormType
            FormType.Text = DocClasses(DocNum).FormTypes(f).Name
                    
            
            Set Pages = ThereXML.createElement("Pages")
            Set Pages = Document.appendChild(Pages)
        
            'SamplesArray = GetSamples(subnode.Attributes.getNamedItem("DocumentClassName").nodeValue)
            For s = 0 To UBound(DocClasses(DocNum).FormTypes(f).Samples)
                Dim Sample As String
                Sample = DocClasses(DocNum).FormTypes(f).Samples(s)
                If Sample <> "" Then
                    'per sample page
                    Set Page = ThereXML.createElement("Page")
                    Set Page = Pages.appendChild(Page)
                    
                    Set ImportFileName = ThereXML.createAttribute("ImportFileName")
                    Page.Attributes.setNamedItem ImportFileName
                    ImportFileName.Text = SamplesDir & "\Samples\" & Sample
                    
                    '*Print #ThereFile, "file" & SampleFileNumber & "=C:\ACXMLAID\Samples\" & Sample
                    'Print #HereFile, "file" & SampleFileNumber & "=" & Source & "-Samples\Samples\" & SamplesArray(i)
                    SampleFileNumber = SampleFileNumber + 1
                End If
            Next s
        Next f
        
    Next
    
    'Close #HereFile
    '*Close #ThereFile
    On Error Resume Next
    fso.CreateFolder XMLDir
    On Error GoTo 0
    
    ThereXML.save XMLDir & "\" & BCFileName & ".xml"
    
Next




'move samples from temp to source dir
If fso.FolderExists(TempPath & "\cabexport\Samples") Then
    fso.MoveFolder TempPath & "\cabexport\Samples", SamplesDir & "\" ' thinks it is a file if no slash
Else
    CabLog "No sample pages present in this cab."
    Log "No sample pages present in this cab."
End If
End Sub





'http://www.codeguru.com/vb/gen/vb_misc/tips/article.php/c2735/
Public Function GetParam(Count As Integer) As String

    Dim i As Long
    Dim j As Integer
    Dim c As String
    Dim bInside As Boolean
    Dim bQuoted As Boolean

    j = 1
    bInside = False
    bQuoted = False
    GetParam = ""

    For i = 1 To Len(Command)

        c = Mid$(Command, i, 1)

        If bInside And bQuoted Then
            If c = """" Then
                j = j + 1
                bInside = False
                bQuoted = False
            End If
        ElseIf bInside And Not bQuoted Then
            If c = " " Then
                j = j + 1
                bInside = False
                bQuoted = False
            End If
        Else
            If c = """" Then
                If j > Count Then Exit Function
                bInside = True
                bQuoted = True
            ElseIf c <> " " Then
                If j > Count Then Exit Function
                bInside = True
                bQuoted = False
            End If
        End If

        If bInside And j = Count And c <> """" Then _
           GetParam = GetParam & c

    Next i

End Function

'http://www.codeguru.com/vb/gen/vb_misc/tips/article.php/c2735/
Public Function GetParamCount() As Integer

    Dim i As Long
    Dim c As String
    Dim bInside As Boolean
    Dim bQuoted As Boolean

    GetParamCount = 0
    bInside = False
    bQuoted = False

    For i = 1 To Len(Command)

        c = Mid$(Command, i, 1)

        If bInside And bQuoted Then
            If c = """" Then
                GetParamCount = GetParamCount + 1
                bInside = False
                bQuoted = False
            End If
        ElseIf bInside And Not bQuoted Then
            If c = " " Then
                GetParamCount = GetParamCount + 1
                bInside = False
                bQuoted = False
            End If
        Else
            If c = """" Then
                bInside = True
                bQuoted = True
            ElseIf c <> " " Then
                bInside = True
                bQuoted = False
            End If
        End If

    Next i

    If bInside Then GetParamCount = GetParamCount + 1

End Function



