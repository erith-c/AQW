# Excel VBA Data #

## APC_QW Class ##

``` VBA
' Class Module: APC_QW
Enum Flag
 Master = 0
 Commercial = 1
 Contractor = 2
End Enum

Const pRelease As String = "release\"
Const pTemplate As String = "template\"
Const pWorking As String = "working\"
Const pReleaseExt As String = ".xlsx"
Const pTemplateExt As String = ".xltx"
Const pPathPrefix As String = "U:\Finance Operations\__PROJECTS__\qw-template\"

Dim pPathPostfix, _
    pVersion, _
    pDivCode, _
    pReleasePath, _
    pTemplatePath, _
    pWorkingPath As String

Dim pManagementFeePercent As Double

Dim pLaborTypes(), _
    pLaborRates(), _
    pSendTo() As Variant
    
Property Get ReleaseFldr() As String
    ReleaseFldr = pRelease
End Property

Property Get TemplateFldr() As String
    TemplateFldr = pTemplate
End Property

Property Get WorkingFldr() As String
    WorkingFldr = pWorking
End Property

Property Get ReleaseExt() As String
    ReleaseExt = pReleaseExt
End Property

Property Get TemplateExt() As String
    TemplateExt = pTemplateExt
End Property

Property Get PathPrefix() As String
    PathPrefix = pPathPrefix
End Property

Property Get PathPostfix() As String
    PathPostfix = pPathPostfix
End Property

Property Let PathPostfix(Path As String)
    pPathPostfix = Path
End Property

Property Get Version() As String
    Version = pVersion
End Property

Property Let Version(Vers As String)
    pVersion = Vers
End Property

Property Get DivCode() As String
    DivCode = pDivCode
End Property

Property Let DivCode(str As String)
    pDivCode = str
End Property

Property Get ManagementFeePercent() As Double
    ManFeePerc = pManagementFeePercent
End Property

Property Let ManagementFeePercent(perc As Double)
    pManagementFeePercent = perc
End Property

Property Get ReleasePath() As String
    ReleasePath = pReleasePath
End Property

Property Let ReleasePath(Path As String)
    pReleasePath = Path
End Property

Property Get TemplatePath() As String
    TemplatePath = pTemplatePath
End Property

Property Let TemplatePath(Path As String)
    pTemplatePath = Path
End Property

Property Get WorkingPath() As String
    WorkingPath = pWorkingPath
End Property

Property Let WorkingPath(Path As String)
    pWorkingPath = Path
End Property

Function GetLaborTypes() As Variant
    GetLaborTypes = pLaborTypes
End Function

Function SetLaborTypes(arr As Variant)
    pLaborTypes = arr
End Function

Function GetLaborRates() As Variant
    GetLaborRates = pLaborRates
End Function

Function SetLaborRates(arr As Variant)
    pLaborRates = arr
End Function
Function GetSendTo() As Variant
    GetSendTo = pSendTo
End Function

Function SetSendTo(arr As Variant)
    pSendTo = arr
End Function
```

## APC_Functions class ##
``` VBA
' Class Module: APC_Functions

' Utility Functions
'# TODO: Export Utility functions to a Utility Class Module

Public Function ArrayToRange(arr As Variant, rng As Range)
' Breaks apart a given Array and puts values into a corresponding Range.
    Dim index As Integer
    Dim cell As Range
    
    index = 0
    
    For Each cell In rng
        cell.value = arr(index)
        index = index + 1
    Next cell
    
End Function

Public Function OpenWB(ByVal Path As String) As Workbook '// Utility function
''' Checks if a file selected by the given path is open.
''' If it is, returns the workbook to the caller.
''' If it isn't, opens the workbook using the given path & returns it to the caller
    
    Dim File As String
    Dim WB As Workbook
    
    File = Dir(Path)
    
    On Error Resume Next
        Set WB = Workbooks(File)
        
        If WB Is Nothing Then
            Set WB = Application.Workbooks.Open(Path)
        End If
    On Error GoTo 0
    
    Set OpenWB = WB

End Function

Public Function Run(FilePath As String)
' Handles the operations of the QW Template creations
    Dim QW As New APC_QW
    Dim WB As Workbook
    
    Set WB = OpenWB(FilePath)
    
    RunCommercial QW, WB
'    RestoreWB WB, 1
'
'    RunContractor QW, WB
'    RestoreWB WB, 2
End Function

' APC QW Functions

Function Increment_Maint(QW As APC_QW, WB As Workbook, Version As Flag)
' Increments the Maintenance Version by one and saves the workbook as a new file.
    Dim Major, Minor, Maint As Range
    Select Case Version
        Case 0
            Set Major = WB.Worksheets("Data").Range("J4")
            Set Minor = WB.Worksheets("Data").Range("J5")
            Set Maint = WB.Worksheets("Data").Range("J6")
        Case 1
            Set Major = WB.Worksheets("Data").Range("J7")
            Set Minor = WB.Worksheets("Data").Range("J8")
            Set Maint = WB.Worksheets("Data").Range("J9")
        Case 2
            Set Major = WB.Worksheets("Data").Range("J10")
            Set Minor = WB.Worksheets("Data").Range("J11")
            Set Maint = WB.Worksheets("Data").Range("J12")
    End Select
    
    Maint.value = Maint.value + 1
    
    Save QW, WB
    QW.Version = Major & "." & Minor & "." & Maint

End Function

Function Prompt(QW As APC_QW)
'Debugging function
    MsgBox "Version: " & QW.Version & vbNewLine _
        & "    Version Code: " & QW.DivCode & vbNewLine & vbNewLine _
        & "Release Path: " & vbNewLine & QW.ReleasePath & vbNewLine & vbNewLine _
        & "Template Path: " & vbNewLine & QW.TemplatePath
End Function

Function Save(QW As APC_QW, WB As Workbook)
' Saves a copy of the workbook
    WB.SaveCopyAs FileName:=QW.ReleasePath
End Function

Function SaveTemplate(QW As APC_QW)
' Opens and saves a file as a Template
    Dim WB As Workbook
    Set WB = Application.Workbooks.Open(FileName:=QW.ReleasePath)
    WB.SaveAs FileName:=QW.TemplatePath, FileFormat:=xlOpenXMLTemplate
    WB.Close
End Function

'' Master Functions

Function Initialize(QW As APC_QW, WB As Workbook, Version As Flag) As Workbook
' Sets the default values for the QW.
'
    Dim Major, Minor, Maint As String
    
    Select Case Version
        Case 0
            WB.Worksheets("Data").Range("A1").value = "Master"
        ''' // Sets the version numbers to the current Master version listed on the Data tab.
            Major = WB.Worksheets("Data").Range("J4").value
            Minor = WB.Worksheets("Data").Range("J5").value
            Maint = WB.Worksheets("Data").Range("J6").value
        ''' // Sets the initial values for the class.
            QW.PathPostfix = "ver_" & Major & "\v_" & Major & "." & Minor & "\QWTemplate_v"
            QW.Version = Major & "." & Minor & "." & Maint
            QW.DivCode = "(master)"
            QW.ManagementFeePercent = 0.16
            QW.SetLaborRates Array("")
            QW.SetLaborTypes Array("")
            QW.SetSendTo Array("")
            QW.ReleasePath = pPathPrefix & pRelease & pPathPostfix & pVersion & pDivCode & pReleaseExt
            QW.TemplatePath = pPathPrefix & pTemplate & pPathPostfix & pVersion & pDivCode & pTemplateExt
            QW.WorkingPath = pPathPrefix & pWorking & pPathPostfix & pVersion & pDivCode & pReleaseExt
        Case 1
            WB.Worksheets("Data").Range("A1").value = "Commercial"
        ''' // Sets the version numbers to the current Commercial version listed on the Data tab.
            Major = WB.Worksheets("Data").Range("J7").value
            Minor = WB.Worksheets("Data").Range("J8").value
            Maint = WB.Worksheets("Data").Range("J9").value
        ''' // Sets the initial values for the class.
            QW.PathPostfix = "ver_" & Major & "\v_" & Major & "." & Minor & "\QWTemplate_v"
            QW.Version = Major & "." & Minor & "." & Maint
            QW.DivCode = "(comm)"
            QW.ManagementFeePercent = 0.16
            QW.SetLaborTypes Array("Install", "ProMedica", "OC", "RLC", "Supervision", "Drafting", "Engineering", "Travel")
            QW.SetLaborRates Array(85#, 70#, 80#, 55#, 100#, 70#, 90#, 49#)
            QW.SetSendTo Array("Laibe Electric", "404 N. Byrne Rd", "Toledo, OH 43606")
            QW.ReleasePath = QW.PathPrefix & QW.ReleaseFldr & QW.PathPostfix & QW.Version & QW.DivCode & QW.ReleaseExt
            QW.TemplatePath = QW.PathPrefix & QW.TemplateFldr & QW.PathPostfix & QW.Version & QW.DivCode & QW.TemplateExt
            QW.WorkingPath = QW.PathPrefix & QW.WorkingFldr & QW.PathPostfix & QW.Version & QW.DivCode & QW.ReleaseExt
        Case 2
            WB.Worksheets("Data").Range("A1").value = "Contractor"
        ''' // Sets the version numbers to the current Contractor version listed on the Data tab.
            Major = WB.Worksheets("Data").Range("J10").value
            Minor = WB.Worksheets("Data").Range("J11").value
            Maint = WB.Worksheets("Data").Range("J12").value
        ''' // Sets the initial values for the class.
            QW.PathPostfix = "ver_" & Major & "\v_" & Major & "." & Minor & "\QWTemplate_v"
            QW.Version = Major & "." & Minor & "." & Maint
            QW.DivCode = "(cont)"
            QW.ManagementFeePercent = 0.175
            QW.SetLaborTypes Array("Install", "Supervision", "Drafting", "Fabrication", "Program", "Travel", "Inspection", "")
            QW.SetLaborRates Array("", "", "", "", "", "", "", "")
            QW.SetSendTo Array("", "", "")
            QW.ReleasePath = pPathPrefix & pRelease & pPathPostfix & pVersion & pDivCode & pReleaseExt
            QW.TemplatePath = pPathPrefix & pTemplate & pPathPostfix & pVersion & pDivCode & pTemplateExt
            QW.WorkingPath = pPathPrefix & pWorking & pPathPostfix & pVersion & pDivCode & pReleaseExt
    End Select

End Function

Function Prep_Master(QW As APC_QW, WB As Workbook)
' Sets workbook to Master functionality.
    Dim sh_QW, sh_PS, sh_Data As Worksheet
    QW.SetLaborTypes Array("", "", "", "", "", "", "", "", "", "", "", "", "")
    QW.SetLaborRates Array("", "", "", "", "", "", "", "", "", "", "", "", "")
    QW.SetSendTo Array("", "", "")

    sh_Data.Range("A1").value = "Master"
    
    ArrayToRange QW.GetLaborTypes, sh_QW.Range("L10:L22")
    ArrayToRange QW.GetLaborRates, sh_QW.Range("N10:N22")
    ArrayToRange QW.GetSendTo, sh_PS.Range("I11:I13")
    
    Increment_Maint QW, WB, 0

    sh_QW.Select
    
    End Function

Function ProtectWB(WB As Workbook, Version As Flag)
' Unprotects & restores visibility for prepping WB
    Dim sh_QW, sh_PM, sh_Trans, sh_MatList, sh_PS, sh_BI, sh_Data, sh_EC As Worksheet
    Dim pw As String
    
    Set sh_QW = WB.Worksheets("Quotation Worksheet")
    Set sh_PM = WB.Worksheets("ProMedica")
    Set sh_Trans = WB.Worksheets("Transmittal")
    Set sh_MatList = WB.Worksheets("Materials List")
    Set sh_PS = WB.Worksheets("Packing Slip")
    Set sh_BI = WB.Worksheets("Billing Insert")
    Set sh_Data = WB.Worksheets("Data")
    Set sh_EC = WB.Worksheets("Error Codes")
    pw = "apc3400!"
    
    Select Case Version
        Case 0       ' Should never run, Master
            GoTo EndOfFunction
        Case 1       ' Commercial
            With sh_PM
                .Select
                .Range("B4:C4").Select
                .Protect Password:=pw
            End With
            With sh_Trans
                .Select
                .Range("E14:W14").Select
                .Protect Password:=pw
            End With
            With sh_MatList
                .Select
                .Range("C6:E6").Select
                .Protect Password:=pw
            End With
            With sh_PS
                .Select
                .Range("I11:K11").Select
                .Protect Password:=pw
            End With
            With sh_BI
                .Select
                .Range("C6:G6").Select
            End With
            sh_Data.Visible = False
            sh_EC.Visible = False
            With sh_QW
                .Select
                .Range("C6:G6").Select
                .Protect Password:=pw
            End With
            GoTo ProtectStructure
        Case 2       ' Contractor
            sh_PM.Visible = False
            With sh_Trans
                .Select
                .Range("E14:W14").Select
                .Protect Password:=pw
            End With
            sh_MatList.Visible = False
            With sh_PS
                .Select
                .Range("I11:K11").Select
                .Protect Password:=pw
            End With
            With sh_BI
                .Select
                .Range("C6:G6").Select
            End With
            sh_Data.Visible = False
            sh_EC.Visible = False
            With sh_QW
                .Select
                .Range("C6:G6").Select
                .Protect Password:=pw
            End With
            GoTo ProtectStructure
    End Select
    
ProtectStructure:
    WB.Protect Password:=pw, Structure:=True
    
EndOfFunction:
End Function


Function RestoreWB(WB As Workbook, Flag As Integer)
' Unprotects & restores visibility for prepping WB
    Dim sh_QW, sh_PM, sh_Trans, sh_MatList, sh_PS, sh_BI, sh_Data, sh_EC As Worksheet
    Dim pw As String
    
    Set sh_QW = WB.Worksheets("Quotation Worksheet")
    Set sh_PM = WB.Worksheets("ProMedica")
    Set sh_Trans = WB.Worksheets("Transmittal")
    Set sh_MatList = WB.Worksheets("Materials List")
    Set sh_PS = WB.Worksheets("Packing Slip")
    Set sh_BI = WB.Worksheets("Billing Insert")
    Set sh_Data = WB.Worksheets("Data")
    Set sh_EC = WB.Worksheets("Error Codes")
    
    pw = "apc3400!"
    WB.Unprotect Password:=pw
    
    Select Case Flag
        Case 0       ' Master
            GoTo EndOfFunction
        Case 1       ' Commercial
            sh_PM.Unprotect Password:=pw
            sh_Trans.Unprotect Password:=pw
            sh_MatList.Unprotect Password:=pw
            sh_PS.Unprotect Password:=pw
            sh_QW.Unprotect Password:=pw
            sh_Data.Visible = True
            sh_EC.Visible = True
        Case 2       ' Contractor
            sh_PM.Visible = True
            sh_Trans.Unprotect Password:=pw
            sh_MatList.Visible = True
            sh_PS.Unprotect Password:=pw
            sh_QW.Unprotect Password:=pw
            sh_Data.Visible = True
            sh_EC.Visible = True
    End Select
    
EndOfFunction:
End Function

'' Commercial Functions

Function Prep_Commercial(QW As APC_QW, Version As Flag)
' Preps the workbook for saving as a Commercial template
    Dim WB As Workbook
    Set WB = OpenWB(QW.ReleasePath)
    
    Dim sh_QW, sh_PS As Worksheet
    Set sh_QW = WB.Worksheets("Quotation Worksheet")
    Set sh_PS = WB.Worksheets("Packing Slip")
    
    ArrayToRange QW.GetLaborTypes, sh_QW.Range("L10:L17")
    ArrayToRange QW.GetLaborRates, sh_QW.Range("N10:N17")
    ArrayToRange QW.GetSendTo, sh_PS.Range("I11:I13")
    
    Increment_Maint WB, QW, Version
    
    ProtectWB WB, Version
    
    WB.Save
    WB.Close
    
End Function

Function Prep_Contractor(QW As APC_QW)
' Preps the workbook for saving as a Commercial template
    Dim WB As Workbook
    Set WB = Application.Workbooks.Open(FileName:=QW.ReleasePath)
    
    Dim pw As String
    pw = "apc3400!"
    
    Dim sh_QW, sh_PM, sh_Trans, sh_MatList, sh_PS, sh_BI, sh_Data, sh_EC As Worksheet
    Set sh_QW = WB.Worksheets("Quotation Worksheet")
    Set sh_PM = WB.Worksheets("ProMedica")
    Set sh_Trans = WB.Worksheets("Transmittal")
    Set sh_MatList = WB.Worksheets("Materials List")
    Set sh_PS = WB.Worksheets("Packing Slip")
    Set sh_BI = WB.Worksheets("Billing Insert")
    Set sh_Data = WB.Worksheets("Data")
    Set sh_EC = WB.Worksheets("Error Codes")
    
    sh_Data.Range("A1").value = "Contractor"
    ArrayToRange QW.GetLaborTypes, sh_QW.Range("L10:L17")
    ArrayToRange QW.GetLaborRates, sh_QW.Range("N10:N17")
    ArrayToRange QW.GetSendTo, sh_PS.Range("I11:I13")
    
    With sh_Trans
        .Select
        .Range("E14:W14").Select
        .Protect Password:=pw
    End With

    With sh_PS
        .Select
        .Range("I11:K11").Select
        .Protect Password:=pw
    End With

    With sh_BI
        .Select
        .Range("C6:G6").Select
    End With
    
    sh_PM.Visible = False
    sh_MatList.Visible = False
    sh_Data.Visible = False
    sh_EC.Visible = False

    With sh_QW
        .Select
        .Range("C6:G6").Select
        .Protect Password:=pw
    End With
    
    WB.Protect Password:=pw, Structure:=True
    WB.Save
    WB.Close
    
End Function

Function RunCommercial(QW As APC_QW, WB As Workbook)
' Handles the operations for creating the Commercial template
    Dim Comm As Flag
    Comm = Flag.Commercial
    Initialize QW, WB, Comm
    Save QW, WB
    Prep_Commercial QW, Comm
    SaveTemplate QW
    
End Function

Function RunContractor(QW As APC_QW, WB As Workbook)
' Handles the operations for creating the Commercial template
     
End Function
```
