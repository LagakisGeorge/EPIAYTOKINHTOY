VERSION 5.00
Begin VB.Form frmDTSTransferObjectsTask 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DTSTransferObjectsTask Object"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "≈Õ«Ã≈—Ÿ”« Ã≈ ’–œÀœ…–¡"
      Height          =   360
      Left            =   4680
      TabIndex        =   5
      Top             =   6360
      Width           =   2895
   End
   Begin VB.CommandButton NeoiPel 
      BackColor       =   &H00FFFF80&
      Caption         =   "Õ›ÔÈ –ÂÎ‹ÙÂÚ-TIMOÀO√…¡"
      Height          =   435
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   3495
   End
   Begin VB.ListBox List2 
      Height          =   5520
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   7575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "¡ÔÛÙÔÎﬁ ”ı„ÍÂÌÙÒ˘ÙÈÍÔı ÛÂ ÷ÔÒÙÁ„¸"
      Height          =   435
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   285
      Top             =   3255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C:\MERCVB\LAPTOP2.TXT  „È· 2Ô LAPTOP"
      Height          =   195
      Left            =   960
      TabIndex        =   4
      Top             =   6720
      Width           =   3285
   End
   Begin VB.Label Label1 
      Caption         =   "¬ÔﬁËÂÈ·"
      Height          =   375
      Left            =   7320
      TabIndex        =   2
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "frmDTSTransferObjectsTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright (C) 2000 Microsoft Corporation
'All rights reserved.
'
'THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
'EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
'MERCHANTIBILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.Option Explicit
'
Public oPackage As New dts.Package
Dim gconnect As String


Dim GDB As New ADODB.Connection

Dim GREM As New ADODB.Connection


Dim ftopUSER
Dim ftoppwd
Dim fConnString


Dim fRServer As String
Dim fRUserName As String

Dim fRPassword As String
Dim fTopikos As String

Dim FLINKED As String

Dim FLAPTOP As String



'Private Sub GenericTaskPackage()
'Dim oConnection As dts.Connection
'Dim oStep As dts.Step
'Dim oTask As dts.Task
'Dim oCustomTask As dts.TransferObjectsTask 'TaskObject
'
'On Error GoTo PackageError
'
''Create step, tasks
''Connections are not necessary when transfering objects.
''In general they are needed when pumping data
'Set oStep = oPackage.Steps.New
'Set oTask = oPackage.Tasks.New("DTSTransferObjectsTask")
'Set oCustomTask = oTask.CustomTask
'
''Fail the package on errors
'oPackage.FailOnError = False
'
'With oStep
'  .Name = "GenericPkgStep"
'  .ExecuteInMainThread = True
'End With
'
'With oTask
'  .Name = "GenericPkgTask"
'End With
''Customize the Task Object
'With oCustomTask
'  .Name = "DTSTransferObjectsTask"
'  'SourceServer property specifies the name of the source server
'  .DropDestinationObjectsFirst = True  'False
'
'  .SourceServer = fRServer '"EBB1C388700542C" ' (local)"
'  'SourceUseTrustedConnection property specifies whether the Windows Authentication security mode is to be used
'
'  .SourceUseTrustedConnection = False
'  .SourceLogin = fRUserName '"sa"
'  .SourcePassword = fRPassword '"38983"
'
'
'
'
'  'DestinationServer property specifies the name of the destination server when you transfer SQL Server objects
'  .DestinationServer = fTopikos '"HPPC"
'
'  'SourceDatabase property specifies the name of the source database
'  .SourceDatabase = "MERCURY"
'
'
'  'DestinationUseTrustedConnection property specifies whether Windows Authentication is used
'  .DestinationUseTrustedConnection = True
'  'DestinationDatabase property specifies the name of the destination database to use
'  .DestinationDatabase = "MERCFORHTOY"
'
'
'  'ScriptFileDirectory property specifies the directory to which the script file and log files are written
'  '.ScriptFileDirectory = 'path must exist
'  'CopyAllObjects property specifies whether to transfer all objects
'  .CopyAllObjects = False
'  'IncludeDependencies property specifies whether dependent objects are scripted and transferred during a transfer
'  .IncludeDependencies = False
'  'IncludeLogins property specifies whether the logins on the source are scripted and transferred
'  .IncludeLogins = False
'  'IncludeUsers property specifies whether the database users on the source are scripted and transferred
'  .IncludeUsers = False
'  'DropDestinationObjectsFirst property specifies whether to drop objects, if they already exist on the destination
'  .DropDestinationObjectsFirst = True
'
'  'CopySchema property specifies whether database objects are copied
'  .CopySchema = True
'  'CopyData property specifies whether data is copied and whether existing data is replaced or appended to
'  .CopyData = DTSTransfer_ReplaceData  ' DTSTransfer_AppendData  '
'
''  'AddObjectForTransfer method adds an object to the list of to be transferred
'  .AddObjectForTransfer "BARCODES", "dbo", DTSSQLObj_UserTable
'  .AddObjectForTransfer "EID", "dbo", DTSSQLObj_UserTable
'  .AddObjectForTransfer "EGG", "dbo", DTSSQLObj_UserTable
'  .AddObjectForTransfer "PEL", "dbo", DTSSQLObj_UserTable
'  .AddObjectForTransfer "EGGTIM", "dbo", DTSSQLObj_UserTable
'  .AddObjectForTransfer "TIM", "dbo", DTSSQLObj_UserTable
''  .AddObjectForTransfer "TIM", "dbo", DTSSQLObj_View
''  .AddObjectForTransfer "EGGTIM", "dbo", DTSSQLObj_StoredProcedure
'
'End With
'
'oStep.TaskName = oCustomTask.Name
'
''Add the step
'oPackage.Steps.Add oStep
'oPackage.Tasks.Add oTask
'
'
'
''GDB.Open "DSN=TOPIKOS"
'
'''''''On Error Resume Next
'Dim K As Long
'
'
'
'
''GDB.Execute "DELETE FROM EID", K
'
''Run the package and release references.
'oPackage.Execute
'GoTo 15   '2 ENTOLES
'
'
'
'Dim R As New ADODB.Recordset
'
'
''R.Open "SELECT COUNT(*) FROM EID", GDB, adOpenForwardOnly, adLockReadOnly
''If R(0) = 0 Then
''  'ƒ≈Õ ≈√…Õ≈ « ≈Õ«Ã≈—Ÿ”«
''  ' –¡…—ÕŸ ‘¡ ”‘œ…◊≈…¡ ¡–œ ‘œ ¡Õ‘…√—¡÷œ  EIDBAC
''  '”‘¡Ã¡‘Ÿ ‘œ –—œ√—¡ÃÃ¡
''  GDB.Execute "INSERT INTO EID SELECT * FROM EIDBAC"
''  End
''
''
''Else
''
''  '≈√…Õ≈ « ≈Õ«Ã≈—Ÿ”« –¡…—ÕŸ ¡Õ‘…√—¡÷¡ ”‘œ EIDBAC
''  On Error Resume Next
''  GDB.Execute "DROP TABLE EIDBAC"
''  GDB.Execute "SELECT * INTO EIDBAC FROM EID"
''End If
'
'
'15
'
'
'
'Label1.Caption = Time$
'  List1.AddItem Time$
'  If List1.ListCount > 30 Then List1.Clear
'
'
'
'
'
'
'
''Clean up
'Set oCustomTask = Nothing
'Set oTask = Nothing
'Set oStep = Nothing
'oPackage.UnInitialize
'Me.Caption = "OK"
'Exit Sub
'
'PackageError:
'Dim sMsg    As String
'  sMsg = "Package failed error: " & sErrorNumConv(Err.Number) & _
'  vbCrLf & Err.Description & vbCrLf & sAccumStepErrors(oPackage)
'  MsgBox sMsg, vbExclamation, oPackage.Name
'  Resume Next
'End Sub

'Private Function sAccumStepErrors(ByVal oPackage As dts.Package) As String
''Accumulate the step error info into the error message.
'Dim oStep       As dts.Step
'Dim sMessage    As String
'Dim lErrNum     As Long
'Dim sDescr      As String
'Dim sSource     As String
'
''Look for steps that completed and failed.
'For Each oStep In oPackage.Steps
'  If oStep.ExecutionStatus = DTSStepExecStat_Completed Then
'    If oStep.ExecutionResult = DTSStepExecResult_Failure Then
'      'Get the step error information and append it to the message.
'      oStep.GetExecutionErrorInfo lErrNum, sSource, sDescr
'      sMessage = sMessage & vbCrLf & _
'      "Step " & oStep.Name & " failed, error: " & _
'      sErrorNumConv(lErrNum) & vbCrLf & sDescr & vbCrLf
'    End If
'  End If
'Next
'
'sAccumStepErrors = sMessage
'End Function
'
'Private Function sErrorNumConv(ByVal lErrNum As Long) As String
''Convert the error number into readable forms, both hex and decimal for the low-order word.
'  If lErrNum < 65536 And lErrNum > -65536 Then
'    sErrorNumConv = "x" & Hex(lErrNum) & ",  " & CStr(lErrNum)
'  Else
'    sErrorNumConv = "x" & Hex(lErrNum) & ",  x" & _
'    Hex(lErrNum And -65536) & " + " & CStr(lErrNum And 65535)
'  End If
'End Function

Private Sub Command1_Click()

' ENHMERVSH PELATON ME YPOLOIPA STO PEDIO LITRA
'GDB LOCAL CONNECTION
'GREM REMOTE CONNECTION


Dim n, c, c2, c3
GDB.Open gconnect   '  "DSN=FORITO;UID=sa;PWD=sa"
Dim linked_server As String:
linked_server = FLINKED 'local  "[KENTRIKOS\SQLGOGAKIS].MERCURY.dbo"

Dim R As New ADODB.Recordset
List2.AddItem "À¡–‘œ– " + FLAPTOP
Dim GMAGAZ As New ADODB.Connection
List2.AddItem "FLINKED " + FLINKED
On Error Resume Next
GMAGAZ.Open fConnString   'DSN REMOTE SERVER ME DEFAULT DATABASE MERCURY P.X. dsn=delloikias2;uid=sa;pwd=12345678
List2.AddItem "≈ÎÂ„˜ÔÚ 2 OK"
 On Error GoTo 0

GMAGAZ.Execute "UPDATE PEL SET CH1=  CAST(ROUND(ISNULL(TYP,0),2)  AS CHAR(12) ) "
GDB.Execute "UPDATE PEL SET CH1=' ' "
GDB.Execute "UPDATE PEL   SET CH1=(SELECT TOP 1 CH1  FROM " + linked_server + ".PEL GG  WHERE GG.KOD=PEL.KOD)"
















End Sub

Private Sub Command2_Click()

Dim n, c, c2, c3








c = InputBox("‰˘ÛÂ ÙÔÌ ·ÒÈËÏ¸ ÙÔı Ûı„ÍÂÌÙÒ˘ÙÈÍÔ˝")



c2 = InputBox("‰˘ÛÂ ÙÔ ·ËÒÔÈÛÏ· Ù˘Ì ÔÛÔÙﬁÙ˘Ì ÙÔı Ûı„ÍÂÌÙÒ˘ÙÈÍÔ˝ „È· ›ÎÂ„˜Ô")

c3 = "Ù" + Format(c, "000000")



GDB.Open gconnect   '  "DSN=FORITO;UID=sa;PWD=sa"


'Dim linked_server As String: linked_server = "[QUEST-PC\DOYTSIOS].mercury.dbo"
Dim linked_server As String: ' MPAINEI STO Ã≈—ÿ–¡‘« ‘œ’ ÷œ—«‘œ’ 3« ”≈…—¡   –..◊.  [OIKIAS].MERCURY.dbo
linked_server = FLINKED '  "[KENTRIKOS\SQLGOGAKIS].MERCURY.dbo"

Dim R As New ADODB.Recordset

R.Open "select sum(POSO) from EGGTIM WHERE LEFT(ATIM,7)='ı' and day(HME)=DAY(GETDATE())", GREM, adOpenDynamic, adLockOptimistic

If R(0) > 0 Then
   MsgBox "’–¡—◊œ’Õ ”«Ã≈—…Õ¡ ‘…ÃœÀœ√…¡ ”‘œ ÷œ—«‘œ. Œ≈÷œ—‘Ÿ”‘≈ ‘œ ÷œ—«‘œ "
   GDB.Close
   Exit Sub
End If
 List2.AddItem "≈ÎÂ„˜ÔÚ 1 OK"
R.Close











 'R.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'PEL'", GREM, adOpenDynamic, adLockOptimistic


'If R(0) > 0 Then
On Error Resume Next
     GREM.Execute "DROP TABLE PEL"
'End If
'R.Close
List2.AddItem "≈ÎÂ„˜ÔÚ 2 OK"
 On Error GoTo 0
     
GREM.Execute "SELECT *  INTO PEL FROM " + linked_server + ".PEL ", n


If n > 0 Then
 List2.AddItem "≈ÎÂ„˜ÔÚ Õ2 ok" + Chr(13) + "PEL ≈√√—¡÷≈” " + Str(n)
Else
   MsgBox "≈ÎÂ„˜ÔÚ Õ2 À¡»œ”" + Chr(13) + "PEL "
   GREM.Close

   Exit Sub
End If



n = 0


 R.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'EID'", GREM, adOpenDynamic, adLockOptimistic


If R(0) > 0 Then
     GREM.Execute "DROP TABLE EID"
End If
R.Close






'GREM.Execute "DROP TABLE EID"
 GREM.Execute "SELECT *  INTO EID  FROM " + linked_server + ".EID", n


If n > 0 Then
    List2.AddItem "≈ÎÂ„˜ÔÚ Õ3 ok" + Chr(13) + "≈…ƒ« ≈√√—¡÷≈” " + Str(n)
Else
   MsgBox "≈ÎÂ„˜ÔÚ Õ3 À¡»œ”" + Chr(13) + "≈…ƒ« "
   GREM.Close

   Exit Sub
End If




n = 0


On Error Resume Next
GREM.Execute "DROP TABLE BARCODES"


On Error GoTo 0
GREM.Execute "SELECT *  INTO BARCODES  FROM " + linked_server + ".BARCODES", n

If n > 0 Then
    List2.AddItem "≈ÎÂ„˜ÔÚ Õ4 ok" + Chr(13) + "BARCODES ≈√√—¡÷≈” " + Str(n)
Else
   MsgBox "≈ÎÂ„˜ÔÚ Õ4 À¡»œ”" + Chr(13) + "BARCODES "
   GREM.Close

   Exit Sub
End If




n = 0



On Error Resume Next
GREM.Execute "DROP TABLE EGGTIM"
On Error GoTo 0




Dim C35 As String, R35 As New ADODB.Recordset
C35 = " FROM " + linked_server + ".EGGTIM  WHERE ATIM='" + c3 + "'"
R35.Open "select count(*) " + C35, GREM, adOpenDynamic
 List2.AddItem "≈ÎÂ„˜ÔÚ Õ5 ok" + Chr(13) + Str(R35(0)) + "EGGTIM ≈√√—¡÷≈” "
R35.Close



GREM.Execute "SELECT *  INTO EGGTIM  " + C35, n

If n > 0 Then
    List2.AddItem "≈ÎÂ„˜ÔÚ Õ5 ok" + Chr(13) + "EGGTIM ≈√√—¡÷≈” " + Str(n)
     GREM.Execute "alter table EGGTIM  drop COLUMN ID"
     GREM.Execute "UPDATE EGGTIM SET ID_NUM=1" ' ≈–≈…ƒ« »¡ Œ¡Õ¡ƒ«Ã…œ’—√«»≈… ‘œ ‘…Ã  ¡… »¡ –¡—≈… ‘œ 1 ‘œ ”’√ ≈Õ‘—Ÿ‘… œ
Else
   MsgBox "≈ÎÂ„˜ÔÚ Õ5 À¡»œ”" + Chr(13) + "EGGTIM ƒ≈Õ Ã≈‘¡÷≈—»« ≈ ‘œ ”’√ ≈Õ‘—Ÿ‘… œ"
  
   
   
   
   GREM.Close
   Exit Sub
End If
n = 0





On Error Resume Next
GREM.Execute "DROP TABLE TIM"
On Error GoTo 0



GREM.Execute "SELECT *  INTO TIM   FROM " + linked_server + ".TIM WHERE ATIM='" + c3 + "'", n
If n > 0 Then
    List2.AddItem "≈ÎÂ„˜ÔÚ Õ6 ok" + Chr(13) + "TIM ≈√√—¡÷≈” " + Str(n)
    GREM.Execute "alter table TIM  drop COLUMN ID_NUM"

Else
   MsgBox "≈ÎÂ„˜ÔÚ Õ6 À¡»œ”" + Chr(13) + "TIM "
   GREM.Close
   Exit Sub
End If
n = 0
GREM.Execute "DROP TABLE EGG"
GREM.Execute "SELECT *  INTO EGG   FROM " + linked_server + ".EGG", n
If n > 0 Then
    List2.AddItem "≈ÎÂ„˜ÔÚ Õ7 ok" + Chr(13) + "EGG ≈√√—¡÷≈” " + Str(n)
     GREM.Execute "alter table EGG  drop COLUMN ID"
Else
   MsgBox "≈ÎÂ„˜ÔÚ Õ7 À¡»œ”" + Chr(13) + "EGG "
   GREM.Close
   Exit Sub
End If
n = 0


'√…¡ Õ¡  ¡Õ≈… UPDATE THN BASH
GREM.Execute "UPDATE PARAMETROI SET TIMH='' WHERE FORMA='MDIFORM1' AND VAR='F_VER' "


GREM.Execute "UPDATE EGG SET IDTIM=999", n







'Dim R As New ADODB.Recordset
R.Open "select sum(POSO) from EGGTIM WHERE LEFT(ATIM,7)='" + c3 + "'", GREM, adOpenDynamic, adLockOptimistic

If R(0) = Val(c2) Then
   List2.AddItem "≈ÎÂ„˜ÔÚ ‘≈À… œ”  ok EGGTIM"
   GREM.Execute "UPDATE EGGTIM SET XRE=POSO,PIS=0"
Else
   MsgBox "≈ÎÂ„˜ÔÚ ‘≈À… œ” À¡»œ”" + Chr(13) + "EGGTIM"
   GREM.Execute "UPDATE EGGTIM SET XRE=POSO,PIS=0"
   GDB.Close

   Exit Sub
End If

R.Close






'Dim r As New ADODB.Recordset
R.Open "select sum(POSO) from EGGTIM WHERE LEFT(ATIM,7)='" + c3 + "'", GREM, adOpenDynamic, adLockOptimistic
If R(0) = Val(c2) Then
   List2.AddItem "≈ÎÂ„˜ÔÚ Õ1 ok" + Chr(13) + "MERCFORHTOY.EGGTIM"
Else
   MsgBox "≈ÎÂ„˜ÔÚ Õ1 À¡»œ”" + Chr(13) + "MERCFORHTOY.EGGTIM"
   GREM.Close
   Exit Sub
End If
R.Close

n = 0



MsgBox "OK"




End Sub

Private Sub Form_Load()
 'Init fail
'Open "c:\mercvb\replica.txt" For Input As #1
'Input #1, fRServer
'Input #1, fRUserName
'Input #1, fRPassword
'Input #1, fTopikos
'Input #1, ftopUSER
'Input #1, ftoppwd
'Input #1, fConnString
'Close #1

 
Open "c:\MERCVB\mercpath.txt" For Input As #1
Input #1, fConnString  'DSN REMOTE SERVER ME DEFAULT DATABASE MERCURY P.X. dsn=delloikias2;uid=sa;pwd=12345678
Input #1, gconnect
Input #1, FLINKED  'linked server (sto forhto) pos blepo ton sql toy kentrikoy p.x. [QUEST-PC\DOYTSIOS].mercury.dbo

Close #1
 FLAPTOP = "LAPTOP1"
 If Len(Dir("C:\MERCVB\LAPTOP2.TXT", vbNormal)) > 0 Then
       FLAPTOP = "LAPTOP2"
 End If
 
 
 GREM.Open fConnString
 
 
 
' oPackage.FailOnError = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
'
End Sub


Private Sub Label1_Click()

MsgBox "ÛÙÔ MercPath.txt ÛÙÁÌ Ò˛ÙÁ ÛÂÈÒ‹ „Ò·ˆ˘ dsn=FORITO;uid=sa;pwd=sa „È· Ì· ÛıÌ‰ÂË˛ ÛÙÔ MERCURY ‘œ’ ÷œ—«‘œ’"
MsgBox "”‘«Õ 3« ”≈…—¡ ‘œ’ MERCPATH.TXT √—¡÷Ÿ ‘œ œÕœÃ¡ KENTRIKOY POY EINAI DHLVMENOS STON FORHTO SAN LINKED SERVER  P.X. [KENTRIKOS\SQLGOGAKIS].MERCURY.dbo"




End Sub

Private Sub NeoiPel_Click()


Dim n, c, c2, c3








'c = InputBox("‰˘ÛÂ ÙÔÌ ·ÒÈËÏ¸ ÙÔı Ûı„ÍÂÌÙÒ˘ÙÈÍÔ˝")



'c2 = InputBox("‰˘ÛÂ ÙÔ ·ËÒÔÈÛÏ· Ù˘Ì ÔÛÔÙﬁÙ˘Ì ÙÔı Ûı„ÍÂÌÙÒ˘ÙÈÍÔ˝ „È· ›ÎÂ„˜Ô")

'c3 = "Ù" + Format(c, "000000")



GDB.Open gconnect   '  "DSN=FORITO;UID=sa;PWD=sa"


'Dim linked_server As String: linked_server = "[QUEST-PC\DOYTSIOS].mercury.dbo"
Dim linked_server As String:
linked_server = FLINKED 'local  "[KENTRIKOS\SQLGOGAKIS].MERCURY.dbo"

Dim R As New ADODB.Recordset

'R.Open "select sum(POSO) from EGGTIM WHERE LEFT(ATIM,7)='ı' and day(HME)=DAY(GETDATE())", GREM, adOpenDynamic, adLockOptimistic
'
'If R(0) > 0 Then
'   MsgBox "’–¡—◊œ’Õ ”«Ã≈—…Õ¡ ‘…ÃœÀœ√…¡ ”‘œ ÷œ—«‘œ. Œ≈÷œ—‘Ÿ”‘≈ ‘œ ÷œ—«‘œ "
'   GDB.Close
'   Exit Sub
'End If
' List2.AddItem "≈ÎÂ„˜ÔÚ 1 OK"
'R.Close


List2.AddItem "À¡–‘œ– " + FLAPTOP



Dim GMAGAZ As New ADODB.Connection
'GMAGAZ.Open "DSN=delloikias2;uid=sa;pwd=12345678"


List2.AddItem "FLINKED " + FLINKED



 'R.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'PEL'", GREM, adOpenDynamic, adLockOptimistic



On Error Resume Next
    ' GREM.Execute "DROP TABLE PEL"


GMAGAZ.Open fConnString   'DSN REMOTE SERVER ME DEFAULT DATABASE MERCURY P.X. dsn=delloikias2;uid=sa;pwd=12345678

List2.AddItem "≈ÎÂ„˜ÔÚ 2 OK"
 On Error GoTo 0
     
Dim PelFields As String
PelFields = "EIDOS,EPO,ONO,DIE,POL,THL,EPA,AFM,PEK,AEG,AYP,TYP,XRESYN,KOD,SHM1,SHM2,KART,XREMHN,PISMHN,XRE,PIS,PMXRE,PMPIS,LASTUPDT,PISSYN,ARTIM,SUMTIM,KODGAL,PLAISIO,ARPARAG,HMELHJ,HME_LHJ,TYPOS,XRVMA,DOY,PLAFON,HMERESAPOP,MEMO,NUM1,NUM2,NUM3,NUM4,HM1,HM2,HM3,HM4,HM5,HM6,CH1,CH2,CH3,CH4,CH5,CH6,ENERGOS,NUM5,NUM6,COMB1,COMB2,COMB3,COMB4,COMB5,EMAIL,KINHTO,PVLHTHS,HM7,HM8,HM9,HM10,HM11,ADT,CH7,COUNTRY"
 
 
 '==============  ƒ«Ã…œ’—√Ÿ ‘œ’ –≈À¡‘≈” –œ’ ƒ≈Õ ’–¡—◊œ’Õ ”‘œ À¡–‘œ–
cc = "insert into PEL(" + PelFields + ") SELECT " + PelFields + "  FROM " + linked_server + ".PEL P  WHERE   P.KOD NOT IN (SELECT KOD FROM PEL  ) "
'≈Õ«Ã≈—ŸÕ≈… ‘œ ‘œ–… œ PEL (laptop forthgoy)  Ã≈ ‘œ’” –≈À¡‘≈” –œ’ ‘œ’ À≈…–œ’Õ ¡–œ ‘œ ¡–œÃ¡ —’”Ã≈Õœ PEL (magazioy,oikias)'
'GMAGAZ.Execute "insert into PEL(KOD,EPO)  SELECT  KOD,EPO FROM [OIKIAS].MERCURY.dbo.PEL P WHERE P.KOD NOT IN (SELECT KOD FROM PEL ) "
     
GDB.Execute cc, n


List2.AddItem "≈ÎÂ„˜ÔÚ 2.1 OK"

'================  ≈Õ«Ã≈—ŸÕ≈… ‘œ ¡–œÃ¡ —’”Ã≈Õœ PEL (magazioy,oikias)  apo to  PEL (laptop forthgoy)  Ã≈ ‘œ’” –≈À¡‘≈” –œ’ ‘œ’ À≈…–œ’Õ
'insert into PEL(KOD,EPO)  SELECT  KOD,EPO FROM [OIKIAS].MERCURY.dbo.PEL P WHERE P.KOD NOT IN (SELECT KOD FROM PEL )
     
cc = "insert into PEL(" + PelFields + ") SELECT  " + PelFields + "  FROM " + "[" + FLAPTOP + "].MERCURY.dbo" + ".PEL P WHERE   P.KOD NOT IN (SELECT KOD FROM PEL  where not (KOD IS NULL)  ) "


GREM.Execute cc, n

List2.AddItem "≈ÎÂ„˜ÔÚ 2.2 OK"



If n >= 0 Then
 List2.AddItem "≈ÎÂ„˜ÔÚ Õ2 ok" + Chr(13) + "PEL ≈√√—¡÷≈” " + Str(n)
Else
  ' MsgBox "≈ÎÂ„˜ÔÚ Õ2 À¡»œ”" + Chr(13) + "PEL "
   'GREM.Close

  ' Exit Sub
End If



n = 0



' ”‘≈ÀÕŸ ‘… ≈√√—¡÷≈”


'–—œ’–œ»≈”« ”‘œ À¡–‘œ– Õ¡ √…Õ≈…
'USE MERCURY
'  ALTER TABLE TIM ADD ID2 UNIQUEIDENTIFIER  default NEWID()
'  ALTER TABLE EGGTIM ADD ID2 VARCHAR(16)
' ALTER TABLE EGG ADD ID2 UNIQUEIDENTIFIER  default NEWID()
'ALTER TABLE EGG ADD ID2TIM VARCHAR(16)


' ¡… ”‘«Õ ¬¡”«
'USE MERCURY
'  ALTER TABLE TIM ADD ID2 VARCHAR(16)
'  ALTER TABLE EGGTIM ADD ID2 VARCHAR(16)
'ALTER TABLE EGG ADD ID2 VARCHAR(16)
'ALTER TABLE EGG ADD ID2TIM VARCHAR(16)

Dim eggtimFields, eggFields, timFields As String
timFields = "KPE ,HME ,TRP ,ATIM ,ART ,AJI ,EIDOS ,METAF ,EKPT ,EIDPAR ,FPA1 ,FPA2 ,FPA3 ,FPA4 ,FPA6 ,FPA7 ,FPA8 ,FPA9 ,TYP ,AJ1 ,AJ2 ,AJ3 ,AJ4 ,AJ5 ,AJ6 ,AJ7 ,AJ8 ,AJ9 ,EKPT1 ,EKPT2 ,EKPT3 ,EKPT4 ,EKPT5 ,HME_KATAX ,KERDOS ,SKOPOS ,SXETIKO ,PARAT ,ELGA ,SYNPOS ,SKOPOS2 ,FORTOSH ,PROOR ,AYTOK ,B_C1 ,B_C2 ,B_N1 ,B_N2 ,KR1 ,KR2 ,ATIM2 ,KLEIDI ,PARAKRATISI ,LITRA ,EFK ,ORA ,ENTITYUID ,ENTITYMARK ,ENTITY ,AADEKAU ,AADEFPA ,ENTLINEN ,INCMARK ,APALAGIFPA ,TYPOMENO ,AKYROMENO ,SXETMARK ,C1 ,C2 ,C3 ,NUM1 ,NUM2 ,NUM3 ,C12 ,C13 ,NUM11 ,ID2"

List2.AddItem "≈ÎÂ„˜ÔÚ 3.1 OK"

eggtimFields = "EIDOS,ATIM,POSO,MONA,TIMM,KERDOS,KODE,HME,ERGO,FPA,PROOD,PROOD_AJ,EKPT,KAU_AJIA,MIK_AJIA,ONOMA,MIKTA,KOLA,PELKOD,PROELEYSH,XRE,PIS,APOT,ATIM2,FCURRENCY,EKPT2,MIKTAKILA,XVRA,LITRA,EFK,AJAGOR,AJPOL,ID2,ID2TIM"

eggFields = "EIDOS ,HME ,HME_KATAX ,EID ,APA ,XRE ,AIT ,XPI ,KOD ,PROOD ,SEIR ,ATIM ,XREOSI ,PISTOSI ,ATIM2 ,USERID ,IDGRA  ,AAXREOPIS ,IDEGGSYND ,ID2,ID2TIM"

List2.AddItem "≈ÎÂ„˜ÔÚ 3.2 OK"

GDB.Execute "UPDATE EGGTIM SET ID2TIM=(SELECT ID2 FROM TIM WHERE ID_NUM=EGGTIM.ID_NUM) WHERE ID2TIM IS NULL"

GDB.Execute "UPDATE EGG SET ID2TIM=(SELECT ID2 FROM TIM WHERE ID_NUM=EGG.IDTIM) WHERE IDTIM>0 AND ID2TIM IS NULL"

List2.AddItem "≈ÎÂ„˜ÔÚ 3.3 OK"

' BAZEI STO KENTRIKO(KATASTHMA) TA TIMOLOGIA TOY LAPTOP
cc = "insert into TIM (" + timFields + ") SELECT  " + timFields + "  FROM " + "[" + FLAPTOP + "].MERCURY.dbo" + ".TIM P WHERE  left(ATIM,1) IN (SELECT EIDOS FROM PARASTAT WHERE POL=1 AND AJIA_APOU<>0) AND  P.ID2 NOT IN (SELECT ID2 FROM TIM where not (ID2 IS NULL) ) "


GREM.Execute cc, n

List2.AddItem "≈ÎÂ„˜ÔÚ 3.4 OK"


If n >= 0 Then
 List2.AddItem "≈ÎÂ„˜ÔÚ Õ ok" + Chr(13) + "TIM ≈√√—¡÷≈” " + Str(n)
Else
  
End If




' BAZEI STO KENTRIKO(KATASTHMA) TA TIMOLOGIA TOY LAPTOP
cc = "insert into EGGTIM (ID_NUM," + eggtimFields + ") SELECT  -1," + eggtimFields + "  FROM " + "[" + FLAPTOP + "].MERCURY.dbo" + ".EGGTIM P WHERE   left(ATIM,1) IN (SELECT EIDOS FROM PARASTAT WHERE POL=1 AND AJIA_APOU<>0)  AND  P.ID2 NOT IN (SELECT ID2 FROM EGGTIM  where not (ID2 IS NULL)  ) "


GREM.Execute cc, n

List2.AddItem "≈ÎÂ„˜ÔÚ 3.8 OK"



If n >= 0 Then
 List2.AddItem "≈ÎÂ„˜ÔÚ Õ ok" + Chr(13) + "EGGTIM ≈√√—¡÷≈” " + Str(n)
Else
  
End If




' BAZEI STO KENTRIKO(KATASTHMA) TA TIMOLOGIA TOY LAPTOP
cc = "insert into EGG (IDTIM," + eggFields + ") SELECT  -1," + eggFields + "  FROM " + "[" + FLAPTOP + "].MERCURY.dbo" + ".EGG P WHERE   P.ID2 NOT IN (SELECT ID2 FROM EGG  where not (ID2 IS NULL)  ) "


GREM.Execute cc, n

List2.AddItem "≈ÎÂ„˜ÔÚ 3.9 OK"



If n >= 0 Then
 List2.AddItem "≈ÎÂ„˜ÔÚ Õ ok" + Chr(13) + "EGG ≈√√—¡÷≈” " + Str(n)
Else
  
End If





GREM.Execute "UPDATE EGGTIM SET ID_NUM=(SELECT TOP 1 ID_NUM FROM TIM WHERE ID2=EGGTIM.ID2TIM ) WHERE ID_NUM=-1"

List2.AddItem "≈ÎÂ„˜ÔÚ 3.10 OK"

GREM.Execute "UPDATE EGG SET IDTIM=(SELECT TOP 1 ID_NUM FROM TIM WHERE ID2=EGG.ID2TIM ) WHERE IDTIM=-1 AND (NOT ID2TIM IS NULL) "

List2.AddItem "≈ÎÂ„˜ÔÚ 3.11 OK"



'===========================================================================================



Exit Sub










 R.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'EID'", GREM, adOpenDynamic, adLockOptimistic


If R(0) > 0 Then
     GREM.Execute "DROP TABLE EID"
End If
R.Close






'GREM.Execute "DROP TABLE EID"
 GREM.Execute "SELECT *  INTO EID  FROM " + linked_server + ".EID", n


If n > 0 Then
    List2.AddItem "≈ÎÂ„˜ÔÚ Õ3 ok" + Chr(13) + "≈…ƒ« ≈√√—¡÷≈” " + Str(n)
Else
   MsgBox "≈ÎÂ„˜ÔÚ Õ3 À¡»œ”" + Chr(13) + "≈…ƒ« "
   GREM.Close

   Exit Sub
End If




n = 0


On Error Resume Next
GREM.Execute "DROP TABLE BARCODES"


On Error GoTo 0
GREM.Execute "SELECT *  INTO BARCODES  FROM " + linked_server + ".BARCODES", n

If n > 0 Then
    List2.AddItem "≈ÎÂ„˜ÔÚ Õ4 ok" + Chr(13) + "BARCODES ≈√√—¡÷≈” " + Str(n)
Else
   MsgBox "≈ÎÂ„˜ÔÚ Õ4 À¡»œ”" + Chr(13) + "BARCODES "
   GREM.Close

   Exit Sub
End If




n = 0



On Error Resume Next
GREM.Execute "DROP TABLE EGGTIM"
On Error GoTo 0




Dim C35 As String, R35 As New ADODB.Recordset
C35 = " FROM " + linked_server + ".EGGTIM  WHERE ATIM='" + c3 + "'"
R35.Open "select count(*) " + C35, GREM, adOpenDynamic
 List2.AddItem "≈ÎÂ„˜ÔÚ Õ5 ok" + Chr(13) + Str(R35(0)) + "EGGTIM ≈√√—¡÷≈” "
R35.Close



GREM.Execute "SELECT *  INTO EGGTIM  " + C35, n

If n > 0 Then
    List2.AddItem "≈ÎÂ„˜ÔÚ Õ5 ok" + Chr(13) + "EGGTIM ≈√√—¡÷≈” " + Str(n)
     GREM.Execute "alter table EGGTIM  drop COLUMN ID"
     GREM.Execute "UPDATE EGGTIM SET ID_NUM=1" ' ≈–≈…ƒ« »¡ Œ¡Õ¡ƒ«Ã…œ’—√«»≈… ‘œ ‘…Ã  ¡… »¡ –¡—≈… ‘œ 1 ‘œ ”’√ ≈Õ‘—Ÿ‘… œ
Else
   MsgBox "≈ÎÂ„˜ÔÚ Õ5 À¡»œ”" + Chr(13) + "EGGTIM ƒ≈Õ Ã≈‘¡÷≈—»« ≈ ‘œ ”’√ ≈Õ‘—Ÿ‘… œ"
  
   
   
   
   GREM.Close
   Exit Sub
End If
n = 0





On Error Resume Next
GREM.Execute "DROP TABLE TIM"
On Error GoTo 0



GREM.Execute "SELECT *  INTO TIM   FROM " + linked_server + ".TIM WHERE ATIM='" + c3 + "'", n
If n > 0 Then
    List2.AddItem "≈ÎÂ„˜ÔÚ Õ6 ok" + Chr(13) + "TIM ≈√√—¡÷≈” " + Str(n)
    GREM.Execute "alter table TIM  drop COLUMN ID_NUM"

Else
   MsgBox "≈ÎÂ„˜ÔÚ Õ6 À¡»œ”" + Chr(13) + "TIM "
   GREM.Close
   Exit Sub
End If
n = 0
GREM.Execute "DROP TABLE EGG"
GREM.Execute "SELECT *  INTO EGG   FROM " + linked_server + ".EGG", n
If n > 0 Then
    List2.AddItem "≈ÎÂ„˜ÔÚ Õ7 ok" + Chr(13) + "EGG ≈√√—¡÷≈” " + Str(n)
     GREM.Execute "alter table EGG  drop COLUMN ID"
Else
   MsgBox "≈ÎÂ„˜ÔÚ Õ7 À¡»œ”" + Chr(13) + "EGG "
   GREM.Close
   Exit Sub
End If
n = 0


'√…¡ Õ¡  ¡Õ≈… UPDATE THN BASH
GREM.Execute "UPDATE PARAMETROI SET TIMH='' WHERE FORMA='MDIFORM1' AND VAR='F_VER' "


GREM.Execute "UPDATE EGG SET IDTIM=999", n







'Dim R As New ADODB.Recordset
R.Open "select sum(POSO) from EGGTIM WHERE LEFT(ATIM,7)='" + c3 + "'", GREM, adOpenDynamic, adLockOptimistic

If R(0) = Val(c2) Then
   List2.AddItem "≈ÎÂ„˜ÔÚ ‘≈À… œ”  ok EGGTIM"
   GREM.Execute "UPDATE EGGTIM SET XRE=POSO,PIS=0"
Else
   MsgBox "≈ÎÂ„˜ÔÚ ‘≈À… œ” À¡»œ”" + Chr(13) + "EGGTIM"
   GREM.Execute "UPDATE EGGTIM SET XRE=POSO,PIS=0"
   GDB.Close

   Exit Sub
End If

R.Close






'Dim r As New ADODB.Recordset
R.Open "select sum(POSO) from EGGTIM WHERE LEFT(ATIM,7)='" + c3 + "'", GREM, adOpenDynamic, adLockOptimistic
If R(0) = Val(c2) Then
   List2.AddItem "≈ÎÂ„˜ÔÚ Õ1 ok" + Chr(13) + "MERCFORHTOY.EGGTIM"
Else
   MsgBox "≈ÎÂ„˜ÔÚ Õ1 À¡»œ”" + Chr(13) + "MERCFORHTOY.EGGTIM"
   GREM.Close
   Exit Sub
End If
R.Close

n = 0



MsgBox "OK"






End Sub

Private Sub Timer1_Timer()
'This sample transfers some SQL Server
'objects from the pubs database to the pubs2
'database. 'Objects' are items like stored
'procedures, triggers, etc..

'If Mid$(Time$, 4, 2) = "00" Then
'   List1.AddItem "START " + Time$
'   GenericTaskPackage
'End If



End Sub
