VERSION 5.00
Begin VB.Form frmDTSTransferObjectsTask 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DTSTransferObjectsTask Object"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   5520
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   7455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "Αποστολή Συγκεντρωτικου σε Φορτηγό"
      Height          =   435
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   285
      Top             =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Βοήθεια"
      Height          =   375
      Left            =   7320
      TabIndex        =   2
      Top             =   240
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
''  'ΔΕΝ ΕΓΙΝΕ Η ΕΝΗΜΕΡΩΣΗ
''  ' ΠΑΙΡΝΩ ΤΑ ΣΤΟΙΧΕΙΑ ΑΠΟ ΤΟ ΑΝΤΙΓΡΑΦΟ  EIDBAC
''  'ΣΤΑΜΑΤΩ ΤΟ ΠΡΟΓΡΑΜΜΑ
''  GDB.Execute "INSERT INTO EID SELECT * FROM EIDBAC"
''  End
''
''
''Else
''
''  'ΕΓΙΝΕ Η ΕΝΗΜΕΡΩΣΗ ΠΑΙΡΝΩ ΑΝΤΙΓΡΑΦΑ ΣΤΟ EIDBAC
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
'Execute the TaskPackage
'GenericTaskPackage










End Sub

Private Sub Command2_Click()

Dim n, c, c2, c3








c = InputBox("δωσε τον αριθμό του συγκεντρωτικού")



c2 = InputBox("δωσε το αθροισμα των ποσοτήτων του συγκεντρωτικού για έλεγχο")

c3 = "τ" + Format(c, "000000")



GDB.Open gconnect   '  "DSN=FORITO;UID=sa;PWD=sa"


'Dim linked_server As String: linked_server = "[QUEST-PC\DOYTSIOS].mercury.dbo"
Dim linked_server As String:
linked_server = FLINKED '  "[KENTRIKOS\SQLGOGAKIS].MERCURY.dbo"

Dim R As New ADODB.Recordset

R.Open "select sum(POSO) from EGGTIM WHERE LEFT(ATIM,7)='υ' and day(HME)=DAY(GETDATE())", GREM, adOpenDynamic, adLockOptimistic

If R(0) > 0 Then
   MsgBox "ΥΠΑΡΧΟΥΝ ΣΗΜΕΡΙΝΑ ΤΙΜΟΛΟΓΙΑ ΣΤΟ ΦΟΡΗΤΟ. ΞΕΦΟΡΤΩΣΤΕ ΤΟ ΦΟΡΗΤΟ "
   GDB.Close
   Exit Sub
End If
 List2.AddItem "Ελεγχος 1 OK"
R.Close











 'R.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'PEL'", GREM, adOpenDynamic, adLockOptimistic


'If R(0) > 0 Then
On Error Resume Next
     GREM.Execute "DROP TABLE PEL"
'End If
'R.Close
List2.AddItem "Ελεγχος 2 OK"
 On Error GoTo 0
     
GREM.Execute "SELECT *  INTO PEL FROM " + linked_server + ".PEL ", n


If n > 0 Then
 List2.AddItem "Ελεγχος Ν2 ok" + Chr(13) + "PEL ΕΓΓΡΑΦΕΣ " + Str(n)
Else
   MsgBox "Ελεγχος Ν2 ΛΑΘΟΣ" + Chr(13) + "PEL "
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
    List2.AddItem "Ελεγχος Ν3 ok" + Chr(13) + "ΕΙΔΗ ΕΓΓΡΑΦΕΣ " + Str(n)
Else
   MsgBox "Ελεγχος Ν3 ΛΑΘΟΣ" + Chr(13) + "ΕΙΔΗ "
   GREM.Close

   Exit Sub
End If




n = 0


On Error Resume Next
GREM.Execute "DROP TABLE BARCODES"


On Error GoTo 0
GREM.Execute "SELECT *  INTO BARCODES  FROM " + linked_server + ".BARCODES", n

If n > 0 Then
    List2.AddItem "Ελεγχος Ν4 ok" + Chr(13) + "BARCODES ΕΓΓΡΑΦΕΣ " + Str(n)
Else
   MsgBox "Ελεγχος Ν4 ΛΑΘΟΣ" + Chr(13) + "BARCODES "
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
 List2.AddItem "Ελεγχος Ν5 ok" + Chr(13) + Str(R35(0)) + "EGGTIM ΕΓΓΡΑΦΕΣ "
R35.Close



GREM.Execute "SELECT *  INTO EGGTIM  " + C35, n

If n > 0 Then
    List2.AddItem "Ελεγχος Ν5 ok" + Chr(13) + "EGGTIM ΕΓΓΡΑΦΕΣ " + Str(n)
     GREM.Execute "alter table EGGTIM  drop COLUMN ID"
     GREM.Execute "UPDATE EGGTIM SET ID_NUM=1" ' ΕΠΕΙΔΗ ΘΑ ΞΑΝΑΔΗΜΙΟΥΡΓΗΘΕΙ ΤΟ ΤΙΜ ΚΑΙ ΘΑ ΠΑΡΕΙ ΤΟ 1 ΤΟ ΣΥΓΚΕΝΤΡΩΤΙΚΟ
Else
   MsgBox "Ελεγχος Ν5 ΛΑΘΟΣ" + Chr(13) + "EGGTIM ΔΕΝ ΜΕΤΑΦΕΡΘΗΚΕ ΤΟ ΣΥΓΚΕΝΤΡΩΤΙΚΟ"
  
   
   
   
   GREM.Close
   Exit Sub
End If
n = 0





On Error Resume Next
GREM.Execute "DROP TABLE TIM"
On Error GoTo 0



GREM.Execute "SELECT *  INTO TIM   FROM " + linked_server + ".TIM WHERE ATIM='" + c3 + "'", n
If n > 0 Then
    List2.AddItem "Ελεγχος Ν6 ok" + Chr(13) + "TIM ΕΓΓΡΑΦΕΣ " + Str(n)
    GREM.Execute "alter table TIM  drop COLUMN ID_NUM"

Else
   MsgBox "Ελεγχος Ν6 ΛΑΘΟΣ" + Chr(13) + "TIM "
   GREM.Close
   Exit Sub
End If
n = 0
GREM.Execute "DROP TABLE EGG"
GREM.Execute "SELECT *  INTO EGG   FROM " + linked_server + ".EGG", n
If n > 0 Then
    List2.AddItem "Ελεγχος Ν7 ok" + Chr(13) + "EGG ΕΓΓΡΑΦΕΣ " + Str(n)
     GREM.Execute "alter table EGG  drop COLUMN ID"
Else
   MsgBox "Ελεγχος Ν7 ΛΑΘΟΣ" + Chr(13) + "EGG "
   GREM.Close
   Exit Sub
End If
n = 0


'ΓΙΑ ΝΑ ΚΑΝΕΙ UPDATE THN BASH
GREM.Execute "UPDATE PARAMETROI SET TIMH='' WHERE FORMA='MDIFORM1' AND VAR='F_VER' "


GREM.Execute "UPDATE EGG SET IDTIM=999", n







'Dim R As New ADODB.Recordset
R.Open "select sum(POSO) from EGGTIM WHERE LEFT(ATIM,7)='" + c3 + "'", GREM, adOpenDynamic, adLockOptimistic

If R(0) = Val(c2) Then
   List2.AddItem "Ελεγχος ΤΕΛΙΚΟΣ  ok EGGTIM"
   GREM.Execute "UPDATE EGGTIM SET XRE=POSO,PIS=0"
Else
   MsgBox "Ελεγχος ΤΕΛΙΚΟΣ ΛΑΘΟΣ" + Chr(13) + "EGGTIM"
   GREM.Execute "UPDATE EGGTIM SET XRE=POSO,PIS=0"
   GDB.Close

   Exit Sub
End If

R.Close






'Dim r As New ADODB.Recordset
R.Open "select sum(POSO) from EGGTIM WHERE LEFT(ATIM,7)='" + c3 + "'", GREM, adOpenDynamic, adLockOptimistic
If R(0) = Val(c2) Then
   List2.AddItem "Ελεγχος Ν1 ok" + Chr(13) + "MERCFORHTOY.EGGTIM"
Else
   MsgBox "Ελεγχος Ν1 ΛΑΘΟΣ" + Chr(13) + "MERCFORHTOY.EGGTIM"
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

 
Open "c:\mercpath.txt" For Input As #1
Input #1, fConnString  'REMOTE SERVER ME DEFAULT DATABASE MERCURY
Input #1, gconnect
Input #1, FLINKED

Close #1
 
 
 GREM.Open fConnString
 
 
 
' oPackage.FailOnError = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
'
End Sub


Private Sub Label1_Click()

MsgBox "στο MercPath.txt στην πρώτη σειρά γραφω dsn=FORITO;uid=sa;pwd=sa για να συνδεθώ στο MERCURY ΤΟΥ ΦΟΡΗΤΟΥ"
MsgBox "ΣΤΗΝ 3Η ΣΕΙΡΑ ΤΟΥ MERCPATH.TXT ΓΡΑΦΩ ΤΟ ΟΝΟΜΑ KENTRIKOY POY EINAI DHLVMENOS STON FORHTO SAN LINKED SERVER  P.X. [KENTRIKOS\SQLGOGAKIS].MERCURY.dbo"




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
