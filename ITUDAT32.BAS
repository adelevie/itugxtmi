Attribute VB_Name = "Module1"
#If Win32 Then
  Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
  Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
  Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
#Else
  Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
  Declare Function GetPrivateProfileInt Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Integer, ByVal lpFileName As String) As Integer
  Declare Function WritePrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lplFileName As String) As Integer
#End If
Dim INIFileName$
Const ptbuffSize% = 45
Const gainbuffSize% = 10
Dim Adm$
Dim Sat_Name$
Dim Long_nom!
Dim N_Diag%
Dim Beam_ID$
Dim EMI_RCP$
Dim Polar_Disc$
Dim Reason$
Dim Diag_Type$




Function Getgain!(Section$)
    Entry$ = "gain"
    DefaultString$ = "0"
    Value$ = Space$(gainbuffSize%)
    x% = GetPrivateProfileString(Section$, Entry$, DefaultString$, Value$, gainbuffSize%, INIFileName$)
    Value$ = Left$(Value$, x%)
    Getgain! = Val(Value$)
End Function

Private Sub GetMainInfo()
    Section$ = "Main"
    Entry$ = "adm"
    DefaultString$ = ""
    MaxSize% = 6
    Value$ = Space$(MaxSize%)
    x% = GetPrivateProfileString(Section$, Entry$, DefaultString$, Value$, MaxSize%, INIFileName$)
    Adm$ = Left$(Value$, x%)


    Section$ = "Main"
    Entry$ = "sat_name"
    DefaultString$ = ""
    MaxSize% = 30
    Value$ = Space$(MaxSize%)
    x% = GetPrivateProfileString(Section$, Entry$, DefaultString$, Value$, MaxSize%, INIFileName$)
    Sat_Name$ = Left$(Value$, x%)


    Section$ = "Main"
    Entry$ = "long_nom"
    DefaultString$ = "0"
    MaxSize% = 12
    Value$ = Space$(MaxSize%)
    x% = GetPrivateProfileString(Section$, Entry$, DefaultString$, Value$, MaxSize%, INIFileName$)
    Long_nom! = Val(Left$(Value$, x%))


    Section$ = "Main"
    Entry$ = "n_diag"
    DefaultString$ = "0"
    MaxSize% = 4
    Value$ = Space$(MaxSize%)
    x% = GetPrivateProfileString(Section$, Entry$, DefaultString$, Value$, MaxSize%, INIFileName$)
    N_Diag% = Val(Left$(Value$, x%))


    Section$ = "Header"
    Entry$ = "beam_id"
    DefaultString$ = ""
    MaxSize% = 8
    Value$ = Space$(MaxSize%)
    x% = GetPrivateProfileString(Section$, Entry$, DefaultString$, Value$, MaxSize%, INIFileName$)
    Beam_ID$ = Left$(Value$, x%)


    Section$ = "Header"
    Entry$ = "emi_rcp"
    Rem Either E for down-link or R for up-link
    DefaultString$ = ""
    MaxSize% = 5
    Value$ = Space$(MaxSize%)
    x% = GetPrivateProfileString(Section$, Entry$, DefaultString$, Value$, MaxSize%, INIFileName$)
    EMI_RCP$ = Left$(Value$, x%)


    Section$ = "Header"
    Entry$ = "polar_disc"
    Rem Either C for co-polarization or X for cross-polarization
    DefaultString$ = ""
    MaxSize% = 5
    Value$ = Space$(MaxSize%)
    x% = GetPrivateProfileString(Section$, Entry$, DefaultString$, Value$, MaxSize%, INIFileName$)
    Polar_Disc$ = Left$(Value$, x%)


    Section$ = "Header"
    Entry$ = "reason"
    Rem A for Advance publication
    Rem C for request for Coordination or modification
    Rem E for Article 14 procedure
    Rem N for notification
    DefaultString$ = ""
    MaxSize% = 5
    Value$ = Space$(MaxSize%)
    x% = GetPrivateProfileString(Section$, Entry$, DefaultString$, Value$, MaxSize%, INIFileName$)
    Reason$ = Left$(Value$, x%)


    Section$ = "Header"
    Entry$ = "type"
    DefaultString$ = ""
    MaxSize% = 6
    Value$ = Space$(MaxSize%)
    x% = GetPrivateProfileString(Section$, Entry$, DefaultString$, Value$, MaxSize%, INIFileName$)
    Diag_Type$ = Left$(Value$, x%)

End Sub

Private Function GetNumPts%(Section$)
    Entry$ = "n_point"
    DefaultInt% = 0
    GetNumPts% = GetPrivateProfileInt(Section$, Entry$, DefaultInt%, INIFileName$)
End Function

Private Function getPtValue$(Section$, pt$)
    Value$ = Space$(ptbuffSize%)
    x% = GetPrivateProfileString(Section$, pt$, "", Value$, ptbuffSize%, INIFileName$)
    getPtValue$ = Trim$(Left$(Value$, x%))
End Function

Private Sub ImportBoreSights(MifFile%, MidFile%)
    For BoreSightIndex% = 1 To NumOfBoreSights%()
        Section$ = "B" + Format$(BoreSightIndex%)
        Value$ = getPtValue$(Section$, "p")
        Lat$ = Mid$(Value$, InStr(Value$, ";") + 1)
        Lon$ = Left$(Value$, InStr(Value$, ";") - 1)

        Print #MifFile%, "Point " + Lon$ + " " + Lat$
        Print #MifFile%, "    Symbol (34,128,10)"
        Write #MidFile%, Adm$, Sat_Name$, Long_nom!, N_Diag%, Beam_ID$, EMI_RCP$, Polar_Disc$, Reason$, Diag_Type$, Section$, 1, 0, "BS"
    Next BoreSightIndex%
End Sub

Private Sub ImportContours(MifFile%, MidFile%)
    For CountourIndex% = 1 To NumOfContours%()
        'List1.Clear
        Section$ = "C" + Format$(CountourIndex%)
        Gain! = Getgain!(Section$)
        NumPts% = GetNumPts%(Section$)
        LastValue$ = getPtValue$(Section$, Pit$(NumPts%))
        For PtIndex% = 1 To NumPts%
            Value$ = getPtValue$(Section$, Pit$(PtIndex%))
            If PtIndex% = 1 Then
                If LastValue$ = Value$ Then
                    Print #MifFile%, "Region 1"
                    Print #MifFile%, "  " + Format$(NumPts%)
                    RegionType$ = "CC"
                Else
                    Print #MifFile%, "Pline " + Format$(NumPts%)
                    RegionType$ = "OC"
                End If
            End If
            Lat$ = Mid$(Value$, InStr(Value$, ";") + 1)
            Lon$ = Left$(Value$, InStr(Value$, ";") - 1)
            'List1.AddItem Pit$(PtIndex%) + Chr$(9) + Value$ + Chr$(9) + Lon$ + Chr$(9) + Lat$
            Print #MifFile%, Lon$ + " " + Lat$
        Next PtIndex%
        Text$ = "Name=" + Section$ + Chr$(10)
        Text$ = Text$ + "NumPts=" + Format$(NumPts%) + Chr$(10)
        Text$ = Text$ + "Gain=" + Format(Gain!) + Chr$(10)
        Text$ = Text$ + "Type=" + RegionType$ + Chr$(10)
        'MsgBox Text$
        Write #MidFile%, Adm$, Sat_Name$, Long_nom!, N_Diag%, Beam_ID$, EMI_RCP$, Polar_Disc$, Reason$, Diag_Type$, Section$, NumPts%, Gain!, RegionType$
        Call MakeMifSettings(MifFile%, RegionType$)
    Next CountourIndex%
End Sub

Private Sub MakeMifHead(FileNum%)
    Print #FileNum%, "Version 300"
    Print #FileNum%, "Charset " + Chr$(34) + "WindowsLatin1" + Chr$(34)
    Print #FileNum%, "Delimiter " + Chr$(34) + "," + Chr$(34)
    Print #FileNum%, "CoordSys Earth Projection 1, 0"
    Print #FileNum%, "Columns 13"
    Print #FileNum%, "  Adm Char(3)"
    Print #FileNum%, "  Sat_Name Char(20)"
    Print #FileNum%, "  Long_Nom Decimal(8, 3)"
    Print #FileNum%, "  N_Diag Smallint"
    Print #FileNum%, "  Beam_ID Char(4)"
    Print #FileNum%, "  EMI_RCP Char(1)"
    Print #FileNum%, "  Polar_Disc Char(1)"
    Print #FileNum%, "  Reason Char(1)"
    Print #FileNum%, "  Diag_Type Char(3)"
    Print #FileNum%, "  ObjName Char(4)"
    Print #FileNum%, "  NumPts Smallint"
    Print #FileNum%, "  Gain Decimal(6, 2)"
    Rem Determins if Boresight or Closed or Open Contour
    Print #FileNum%, "  ObjType Char(2)"
    Print #FileNum%, "Data"
    Print #FileNum%, ""
End Sub

Private Sub MakeMifSettings(FileNum%, RegionType$)
    Select Case RegionType$
        Case "CC"
            Print #FileNum%, "    Pen (1,2,128)"
            Print #FileNum%, "    Brush (1,0,16777215)"
        Case "OC"
            Print #FileNum%, "    Pen (1,2,128)"
        Case "BS"
            Print #FileNum%, "    Symbol (34,128,10)"
        Case Else
    End Select
End Sub

Private Function NumOfBoreSights%()
    Section$ = "Boresights"
    Entry$ = "n_bore"
    DefaultInt% = 0
    NumOfBoreSights% = GetPrivateProfileInt(Section$, Entry$, DefaultInt%, INIFileName$)
End Function

Private Function NumOfContours%()
    Section$ = "Contours"
    Entry$ = "n_cont"
    DefaultInt% = 0
    NumOfContours% = GetPrivateProfileInt(Section$, Entry$, DefaultInt%, INIFileName$)
End Function

Private Function Pit$(TempInt%)
    Pit$ = "p" + Format$(TempInt%)
End Function


Sub main()
  INIFileName$ = Command$
  If INIFileName$ <> "" Then
    BaseFileName$ = Left$(INIFileName$, InStr(INIFileName$, "."))
    MidFileName$ = BaseFileName$ + "Mid"
    MidFile% = FreeFile
    Open MidFileName$ For Output As #MidFile%

    MifFileName$ = BaseFileName$ + "Mif"
    MifFile% = FreeFile
    Open MifFileName$ For Output As #MifFile%

    Call GetMainInfo

    Call MakeMifHead(MifFile%)

    Call ImportContours(MifFile%, MidFile%)
    Call ImportBoreSights(MifFile%, MidFile%)


    Close #MidFile%
    Close #MifFile%
    MsgBox INIFileName$ + " inported into " + MifFileName$
  Else
    MsgBox "No filename passed on the command line!" + Chr$(10) + "This program imports ITU GXT data into a MapInfo MIF file format." + Chr$(10) + App.EXEName + " filename.gxt"
  End If
End Sub
