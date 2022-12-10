VERSION 5.00
Begin VB.Form frmDTSTransferObjectsTask 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DTSTransferObjectsTask Object"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   6885
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   7455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "ληψη απο φορτηγό"
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   4575
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   285
      Top             =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Παραλαβή αρχείων από απομακρυσμένο"
      Height          =   315
      Left            =   3600
      TabIndex        =   0
      Top             =   8400
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label1 
      Caption         =   "ΤΙΜΟΛΟΓΙΑ = υ"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   7680
      Width           =   7455
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
'Public oPackage As New DTS.Package


Dim GDB As New ADODB.Connection
Dim GREM As New ADODB.Connection

Dim fConnString As String


Dim fRServer As String
Dim fRUserName As String

Dim fRPassword As String
Dim fTopikos As String





Private Sub GenericTaskPackage()
End Sub

Private Function sAccumStepErrors(ByVal oPackage As DTS.Package) As String
End Function

Private Function sErrorNumConv(ByVal lErrNum As Long) As String
End Function

Private Sub Command1_Click()
'Execute the TaskPackage
'GenericTaskPackage
End Sub

Private Sub Command2_Click()

Dim n, c, c2, c3
'
' c = InputBox("δωσε τον αριθμό του συγκεντρωτικού")
'
'c2 = InputBox("δωσε το αθροισμα των ποσοτήτων του συγκεντρωτικού για έλεγχο")
'
'c3 = "τ" + Format(c, "000000")
'
Open "c:\MERCPATH.TXT" For Input As #1
   Input #1, fConnString
   Input #1, c
   Input #1, c2  ' LINKED O KENTRIKOS (SE FORITO H/Y)
   Input #1, FLINKED
   
Close #1

GREM.Open fConnString

List2.AddItem "ΣΥΝΔΕΣΗ ΜΕ ΦΟΡΗΤΟ ΟΚ"

Dim Linked_Server As String: Linked_Server = FLINKED ' "LENOVO.MERCURY.dbo."

GDB.Open c ' "DSN=FORITO;UID=sa;PWD=sa"
'

List2.AddItem "ΣΥΝΔΕΣΗ ΜΕ ΤΟΠΙΚΟ ΟΚ"




Dim r As New ADODB.Recordset




r.Open "SELECT SUM(POSO) FROM  EGGTIM  WHERE LEFT(ATIM,1)='υ' AND DAY(HME)=DAY(GETDATE()) AND MONTH(HME)=MONTH(GETDATE()) AND YEAR(HME)=YEAR(GETDATE())", GDB, adOpenDynamic, adLockOptimistic

Dim ANS
If r(0) > 0 Then
   ANS = MsgBox("ΠΡΟΣΟΧΗ ΥΠΑΡΧΟΥΝ ΤΑ ΣΗΜΕΡΙΝΑ ΤΙΜΟΛΟΓΙΑ ΣΤΟ ΑΡΧΕΙΟ. ΝΑ ΠΡΟΧΩΡΗΣΩ ; ", vbYesNo)
   If ANS = vbNo Then
      Exit Sub
   End If
   
End If

r.Close

Dim ar_sygk As Long

r.Open "SELECT ARITMISI FROM ARITMISI WHERE ID=32", GDB, adOpenDynamic, adLockOptimistic
ar_sygk = r(0) + 1
r.Close




 GDB.BeginTrans

On Error Resume Next

Kill "c:\MERCvb\sql.TXT"


'GDB.Execute "DROP TABLE TEMP_EGG"
'GDB.Execute "DROP TABLE TEMP_EGGTIM"
'GDB.Execute "DROP TABLE TEMP_TIM"
'GDB.Execute "DROP TABLE TEMP_EGGTIM"
On Error GoTo 0

'εισαγωγη εγγραφων EGG
'c = "INSERT INTO EGG(EIDOS,HME,HME_KATAX,EID,APA,XRE,AIT,XPI,KOD,PROOD,SEIR,ATIM,XREOSI,PISTOSI,ATIM2,USERID,IDGRA,IDTIM )   "
'c = c + " SELECT EIDOS,HME,HME_KATAX,EID,APA,XRE,AIT,XPI,KOD,PROOD,SEIR,ATIM,XREOSI,PISTOSI,ATIM2,USERID,"
'c = c + " IDGRA,IDTIM  FROM LENOVO.mercury.dbo.EGG  f  WHERE   IDTIM<>999"
'c = "SELECT * INTO EGG "
'c = c + "  FROM LENOVO.mercury.dbo.EGG  f  WHERE   IDTIM<>999"  ' f.ID NOT IN (SELECT ID FROM EGG )"
'GDB.Execute c, n



c = "INSERT INTO EGG(EIDOS,HME,HME_KATAX,EID,APA,XRE,AIT,XPI,KOD,PROOD,SEIR,ATIM,XREOSI,PISTOSI,ATIM2,USERID,IDGRA,IDTIM )   "
c = c + " SELECT EIDOS,HME,HME_KATAX,EID,APA,XRE,AIT,XPI,KOD,PROOD,SEIR,ATIM,XREOSI,PISTOSI,ATIM2,USERID,"
c = c + " IDGRA,IDTIM  FROM " + Linked_Server + "EGG  f  WHERE   IDTIM is null"
'c = "SELECT * INTO EGG "
'c = c + "  FROM "+linked_server+"EGG  f  WHERE   IDTIM<>999"  ' f.ID NOT IN (SELECT ID FROM EGG )"
GDB.Execute c, n

Open "c:\MERCvb\sql.TXT" For Append As #5
  Print #5, c
Close #5


If n > 0 Then
   List2.AddItem "Ελεγχος Ν1 ok" + Chr(13) + "EGG ΕΓΓΡΑΦΕΣ " + Str(n)
Else
   MsgBox "Ελεγχος Ν1 ΛΑΘΟΣ" + Chr(13) + "EGG "
   GDB.RollbackTrans
   GDB.Close
   Exit Sub
End If

'GREM.Execute "update EGGTIM SET KOLA=ID_NUM" ' ΚΡΑΤΑΩ ΤΟ ID POY EXOYN GIATI UA XAUEI
'GREM.Execute "update TIM SET KR2=ID_NUM" ' ΚΡΑΤΑΩ ΤΟ ID POY EXOYN GIATI UA XAUEI

List2.AddItem "ΚΡΑΤΩ ID"

'δρισκω ποια παραστατικα εκοψε για να τα βαλω στα σχολια
Dim RTT As New ADODB.Recordset
RTT.Open "SELECT MAX(ATIM) AS C20,MIN(ATIM) AS C10 FROM TIM WHERE LEFT(ATIM,1)='υ' ", GREM, adOpenDynamic, adLockOptimistic

Dim CMAX As String, CMIN As String
CMAX = LTrim(Str(Val(Mid(RTT(0), 2, 6))))
CMIN = LTrim(Str(Val(Mid(RTT(1), 2, 6))))

RTT.Close


List2.AddItem "ΒΡΙΣΚΩ ΑΠΟ ΠΟΙΟ ΑΡΧΙΣΕ ΚΑΙ ΣΕ ΠΟΙΟ ΤΕΛΕΙΩΣΕ"





Dim FEGGTIM As String
FEGGTIM = "EIDOS,ATIM,POSO,MONA,TIMM,KERDOS,KODE,HME,ERGO,FPA,PROOD,PROOD_AJ,EKPT,KAU_AJIA,MIK_AJIA,ONOMA,MIKTA,KOLA,PELKOD,PROELEYSH,XRE,PIS,APOT,ATIM2,FCURRENCY,EKPT2,ID_NUM,MIKTAKILA"



'εισαγωγη εγγραφων EGGTIM ΤΙΜΟΛΟΓΙΩΝ ΗΜΕΡΑΣ
c = "INSERT INTO EGGTIM(" + FEGGTIM + ")  SELECT " + FEGGTIM + "  from " + Linked_Server + "EGGTIM WHERE LEFT(ATIM,1)<>'τ'"
GDB.Execute c, n


Open "c:\MERCvb\sql.TXT" For Append As #5
  Print #5, c
Close #5





If n > 0 Then
   List2.AddItem "Ελεγχος Ν2 ok" + Chr(13) + "EGGTIM ΕΓΓΡΑΦΕΣ " + Str(n)
Else
   MsgBox "Ελεγχος Ν2 ΛΑΘΟΣ" + Chr(13) + "EGGTIM "
   GDB.RollbackTrans
   GDB.Close
   Exit Sub
End If


'εισαγωγη εγγραφων TIM ΤΙΜΟΛΟΓΙΩΝ ΗΜΕΡΑΣ



Dim COLTIM As String
COLTIM = "KPE,HME,TRP,ATIM,ART,AJI,EIDOS,METAF,EKPT,EIDPAR,FPA1,FPA2,FPA3,FPA4,FPA6,FPA7,FPA8,FPA9,TYP,AJ1, AJ2,AJ3,AJ4,AJ5,AJ6,AJ7,AJ8,AJ9,EKPT1,EKPT2,EKPT3,EKPT4,EKPT5,HME_KATAX,KERDOS,SKOPOS,SXETIKO,PARAT,ELGA,SYNPOS,SKOPOS2,FORTOSH,PROOR,AYTOK,B_C1,B_C2,B_N1,B_N2,KR1,KR2,ATIM2"

c = "insert into  TIM(" + COLTIM + ") select " + COLTIM + " from " + Linked_Server + "TIM WHERE LEFT(ATIM,1)<>'τ'"

GDB.Execute c, n

Open "c:\MERCvb\sql.TXT" For Append As #5
  Print #5, c
Close #5



If n > 0 Then
   List2.AddItem "Ελεγχος Ν3 ok" + Chr(13) + "TIM ΕΓΓΡΑΦΕΣ " + Str(n)
Else
   MsgBox "Ελεγχος Ν3 ΛΑΘΟΣ" + Chr(13) + "TIM "
   
   GDB.RollbackTrans
   GDB.Close
   Exit Sub
End If



'------------ DHMIOYRGIA SYGKENTRVTIKOY  ----------------------------

On Error Resume Next
GDB.Execute "DROP TABLE DOKSYGKENTROTIKO"
GDB.Execute " drop TABLE DOKFORTHGO"
On Error GoTo 0
 
 
'==============================================================================================================================

'SOYMES STO DOKFORTHGO
GDB.Execute "SELECT SUM(XRE) AS FORT ,SUM(PIS) AS POL , SUM(XRE)-SUM(PIS) AS YPOL,KODE INTO DOKFORTHGO  FROM " + Linked_Server + "EGGTIM GROUP BY KODE", n
If n > 0 Then
   List2.AddItem "Ελεγχος Ν4 ok" + Chr(13) + "ΥΠΟΛΟΓΙΣΜΟΣ ΠΩΛΗΣΕΩΝ " + Str(n)
Else
   MsgBox "Ελεγχος Ν4 ΛΑΘΟΣ" + Chr(13) + "ΥΠΟΛΟΓΙΣΜΟΣ ΠΩΛΗΣΕΩΝ "
   GDB.RollbackTrans
   GDB.Close
   Exit Sub
End If

'STO DOKSYGKENTROTIKO  BAZO TO PALIO SYGKENTROTIKO

'GDB.Execute "DROP TABLE DOKSYGKENTROTIKO"
GDB.Execute "SELECT  *  INTO DOKSYGKENTROTIKO   FROM " + Linked_Server + "EGGTIM WHERE LEFT(ATIM,1)='τ'", n
If n > 0 Then
   List2.AddItem "Ελεγχος Ν5 ok" + Chr(13) + "ΥΠΟΛΟΓΙΣΜΟΣ ΠΩΛΗΣΕΩΝ " + Str(n)
Else
   MsgBox "Ελεγχος Ν5 ΛΑΘΟΣ" + Chr(13) + "ΥΠΟΛΟΓΙΣΜΟΣ ΠΩΛΗΣΕΩΝ "
   GDB.RollbackTrans
   GDB.Close
   Exit Sub
End If


'ENHMERONO TO SDA  ME TIS POLHSEIS
GDB.Execute "update  DOKSYGKENTROTIKO set MIKTA=DOKFORTHGO.POL    FROM  DOKSYGKENTROTIKO LEFT JOIN DOKFORTHGO ON DOKSYGKENTROTIKO.KODE=DOKFORTHGO.KODE ", n
If n > 0 Then
   List2.AddItem "Ελεγχος Ν6 ok" + Chr(13) + "ΥΠΟΛΟΓΙΣΜΟΣ ΠΩΛΗΣΕΩΝ (UPDATE ΜΙΚΤΑ)" + Str(n)
Else
   MsgBox "Ελεγχος Ν6 ΛΑΘΟΣ" + Chr(13) + "ΥΠΟΛΟΓΙΣΜΟΣ ΠΩΛΗΣΕΩΝ (UPDATE ΜΙΚΤΑ)"
   GDB.RollbackTrans
   GDB.Close
   Exit Sub
End If



'== ενημερωνω το συγκεντρωτικο  με τη σωστη ημερομηνια και αριθμο παραστατικου  (HME,ATIM)

GDB.Execute "UPDATE DOKSYGKENTROTIKO SET ATIM= 'τ" + Format(ar_sygk, "000000") + "'    ", n
If n > 0 Then
   List2.AddItem "Ελεγχος Ν7 ok" + Chr(13) + "ΥΠΟΛΟΓΙΣΜΟΣ ΠΩΛΗΣΕΩΝ (UPDATE ATIM)" + Str(n)
Else
   MsgBox "Ελεγχος Ν7 ΛΑΘΟΣ" + Chr(13) + "ΥΠΟΛΟΓΙΣΜΟΣ ΠΩΛΗΣΕΩΝ (UPDATE ATIM)"
   GDB.RollbackTrans
   GDB.Close
   Exit Sub
End If

GDB.Execute "UPDATE DOKSYGKENTROTIKO  SET HME=GETDATE(),MIKTAKILA=POSO-MIKTA", n
If n > 0 Then
   List2.AddItem "Ελεγχος Ν8 ok" + Chr(13) + "ΥΠΟΛΟΓΙΣΜΟΣ ΠΩΛΗΣΕΩΝ (UPDATE ΜΙΚΤΑ)" + Str(n)
Else
   MsgBox "Ελεγχος Ν8 ΛΑΘΟΣ" + Chr(13) + "ΥΠΟΛΟΓΙΣΜΟΣ ΠΩΛΗΣΕΩΝ (UPDATE ΜΙΚΤΑ)"
   GDB.RollbackTrans
   GDB.Close
   Exit Sub
End If










GDB.Execute "INSERT INTO EGGTIM(" + FEGGTIM + ")  SELECT " + FEGGTIM + " FROM  DOKSYGKENTROTIKO", n
If n > 0 Then
   List2.AddItem "Ελεγχος Ν9 ok" + Chr(13) + "ΕΙΣΑΓΩΓΗ ΣΥΓΚΕΝΤΡΩΤΙΚΟΥ ΠΩΛΗΣΕΩΝ (INSERT INTO EGGTIM)" + Str(n)
Else
   MsgBox "Ελεγχος Ν9 ΛΑΘΟΣ" + Chr(13) + "ΕΙΣΑΓΩΓΗ ΣΥΓΚΕΝΤΡΩΤΙΚΟΥ ΠΩΛΗΣΕΩΝ (INSERT INTO EGGTIM)"
   GDB.RollbackTrans
   GDB.Close
   Exit Sub
End If







'ΔΙΝΩ ΝΕΟ ΑΡΙΘΜΟ ΣΤΟ ΣΥΓΚΕΝΤΡΩΤΙΚΟ ΠΟΥ ΘΑ ΕΙΣΑΧΘΕΙ
GREM.Execute "UPDATE TIM SET  PARAT='TIMOΟΛΓΙΑ ΑΠΟ " + CMIN + " ΕΩΣ ΚΑΙ " + CMAX + " ', ATIM= 'τ" + Format(ar_sygk, "000000") + "' WHERE LEFT(ATIM,1)='τ'   ", n
'
'+RTRIM(RIGHT(CONVERT(VARCHAR(7),1000000+1+CONVERT(INT,SUBSTRING(ATIM,2,6))) ,6))




If n > 0 Then
   List2.AddItem "Ελεγχος Ν10 ok" + Chr(13) + "ΥΠΟΛΟΓΙΣΜΟΣ ΠΩΛΗΣΕΩΝ (UPDATE ATIM)" + Str(n)
Else
   MsgBox "Ελεγχος Ν10 ΛΑΘΟΣ" + Chr(13) + "ΥΠΟΛΟΓΙΣΜΟΣ ΠΩΛΗΣΕΩΝ (UPDATE ATIM)"
   
'   GDB.RollbackTrans
 '  GDB.Close
  ' Exit Sub
End If



'c = "insert into TIM  SELECT KPE,HME,TRP,ATIM,ART,AJI,EIDOS, "
'c = c + " METAF,EKPT,EIDPAR,FPA1,FPA2,FPA3,FPA4,FPA6,FPA7,FPA8,FPA9,TYP,AJ1, "
'c = c + " AJ2,AJ3,AJ4,AJ5,AJ6,AJ7,AJ8,AJ9,EKPT1,EKPT2,EKPT3,EKPT4,EKPT5,HME_KATAX,KERDOS,SKOPOS,SXETIKO,PARAT,ELGA,SYNPOS,SKOPOS2,FORTOSH,PROOR,AYTOK,B_C1,B_C2,B_N1,B_N2,KR1,KR2,ATIM2,IDG"
'c = c + " from " + Linked_Server + "TIM WHERE LEFT(ATIM,1)='τ'"

'GDB.Execute c, n
 
c = "insert into TIM(" + COLTIM + ")  SELECT " + COLTIM + ""
c = c + " from " + Linked_Server + "TIM WHERE LEFT(ATIM,1)='τ'"

GDB.Execute c, n
 
 
 
 
 
 

If n > 0 Then
   List2.AddItem "Ελεγχος Ν11 ok" + Chr(13) + "ΕΙΣΑΓΩΓΗ  (INSERT TIM)" + Str(n)
Else
   MsgBox "Ελεγχος Ν11 ΛΑΘΟΣ" + Chr(13) + "ΕΙΣΑΓΩΓΗ ΠΩΛΗΣΕΩΝ (INSERT TIM)"
   GDB.RollbackTrans
   GDB.Close
   Exit Sub
End If

'   GDB.RollbackTrans











GDB.Execute "update ARITMISI set ARITMISI=ARITMISI+1 where ID=32", n

If n > 0 Then
   List2.AddItem "Ελεγχος Ν12 ok" + Chr(13) + "ΕΙΣΑΓΩΓΗ  (UPDATE ARITMISI )" + Str(n)
Else
   MsgBox "Ελεγχος Ν12 ΛΑΘΟΣ" + Chr(13) + "ΕΙΣΑΓΩΓΗ ΠΩΛΗΣΕΩΝ (UPDATE ARITMISI)"
   GDB.RollbackTrans
   GDB.Close
   Exit Sub
End If









'GREM.Execute "update EGGTIM SET KOLA=ID_NUM" ' ΚΡΑΤΑΩ ΤΟ ID POY EXOYN GIATI UA XAUEI
'GREM.Execute "update TIM SET KR2=ID_NUM" ' ΚΡΑΤΑΩ ΤΟ ID POY EXOYN GIATI UA XAUEI



'GDB.Execute " UPDATE EGGTIM SET ID_NUM=0(SELECT TOP 1 ID_NUM FROM TIM WHERE KR2>0 AND EGGTIM.KOLA=TIM.KR2)"
GDB.Execute " UPDATE EGGTIM SET KOLA=0  WHERE LEFT(ATIM,1)='υ' AND KOLA>0"
GDB.Execute " UPDATE TIM    SET KR2=0   WHERE LEFT(ATIM,1)='υ' AND KR2>0"















GDB.CommitTrans




MsgBox "OK"






















'
'
'r.Open "select sum(POSO) from MERCFORHTOY.dbo.EGGTIM WHERE LEFT(ATIM,7)='" + c3 + "'", GDB, adOpenDynamic, adLockOptimistic
'
'If r(0) = Val(c2) Then
'   List2.AddItem "Ελεγχος Ν1 ok" + Chr(13) + "MERCFORHTOY.EGGTIM"
'Else
'   MsgBox "Ελεγχος Ν1 ΛΑΘΟΣ" + Chr(13) + "MERCFORHTOY.EGGTIM"
'   GDB.Close
'
'   Exit Sub
'End If
'
'r.Close
'
'n = 0
'
'
'
'
'On Error Resume Next
'GDB.Execute "DROP TABLE PEL"
'GDB.Execute "SELECT *  INTO PEL  FROM MERCFORHTOY.dbo.PEL", n
'
'
'If n > 0 Then
' List2.AddItem "Ελεγχος Ν2 ok" + Chr(13) + "PEL ΕΓΓΡΑΦΕΣ " + Str(n)
'Else
'   MsgBox "Ελεγχος Ν2 ΛΑΘΟΣ" + Chr(13) + "PEL "
'   GDB.Close
'
'   Exit Sub
'End If
'
'
'
'n = 0
'
'
'
'GDB.Execute "DROP TABLE EID"
' GDB.Execute "SELECT *  INTO EID  FROM MERCFORHTOY.dbo.EID", n
'
'
'If n > 0 Then
'    List2.AddItem "Ελεγχος Ν3 ok" + Chr(13) + "ΕΙΔΗ ΕΓΓΡΑΦΕΣ " + Str(n)
'Else
'   MsgBox "Ελεγχος Ν3 ΛΑΘΟΣ" + Chr(13) + "ΕΙΔΗ "
'   GDB.Close
'
'   Exit Sub
'End If
'
'
'
'
'n = 0
'
'
'
'GDB.Execute "DROP TABLE BARCODES"
'GDB.Execute "SELECT *  INTO BARCODES  FROM MERCFORHTOY.dbo.BARCODES", n
'
'If n > 0 Then
'    List2.AddItem "Ελεγχος Ν4 ok" + Chr(13) + "BARCODES ΕΓΓΡΑΦΕΣ " + Str(n)
'Else
'   MsgBox "Ελεγχος Ν4 ΛΑΘΟΣ" + Chr(13) + "BARCODES "
'   GDB.Close
'
'   Exit Sub
'End If
'
'
'
'
'n = 0
'
'
'
'
'GDB.Execute "DROP TABLE EGGTIM"
'GDB.Execute "SELECT *  INTO EGGTIM   FROM MERCFORHTOY.dbo.EGGTIM WHERE ATIM='" + c3 + "'", n
'
'If n > 0 Then
'    List2.AddItem "Ελεγχος Ν5 ok" + Chr(13) + "EGGTIM ΕΓΓΡΑΦΕΣ " + Str(n)
'Else
'   MsgBox "Ελεγχος Ν5 ΛΑΘΟΣ" + Chr(13) + "EGGTIM "
'   GDB.Close
'
'   Exit Sub
'End If
'
'
'
'
'n = 0
'
'
'
'GDB.Execute "DROP TABLE TIM"
'GDB.Execute "SELECT *  INTO TIM   FROM MERCFORHTOY.dbo.TIM WHERE ATIM='" + c3 + "'", n
'
'
'If n > 0 Then
'    List2.AddItem "Ελεγχος Ν6 ok" + Chr(13) + "TIM ΕΓΓΡΑΦΕΣ " + Str(n)
'Else
'   MsgBox "Ελεγχος Ν6 ΛΑΘΟΣ" + Chr(13) + "TIM "
'   GDB.Close
'
'   Exit Sub
'End If
'
'
'
'
'n = 0
'
'
'
'
'
'GDB.Execute "DROP TABLE EGG"
'GDB.Execute "SELECT *  INTO EGG   FROM MERCFORHTOY.dbo.EGG", n
'
'If n > 0 Then
'    List2.AddItem "Ελεγχος Ν7 ok" + Chr(13) + "EGG ΕΓΓΡΑΦΕΣ " + Str(n)
'Else
'   MsgBox "Ελεγχος Ν7 ΛΑΘΟΣ" + Chr(13) + "EGG "
'   GDB.Close
'
'   Exit Sub
'End If
'
'
'
'n = 0
'
'
'
'
'
'GDB.Execute "UPDATE EGG SET IDTIM=999", n
'
'
'
'
'
'
'
'r.Open "select sum(POSO) from EGGTIM WHERE LEFT(ATIM,7)='" + c3 + "'", GDB, adOpenDynamic, adLockOptimistic
'
'If r(0) = Val(c2) Then
'   List2.AddItem "Ελεγχος ΤΕΛΙΚΟΣ  ok EGGTIM"
'Else
'   MsgBox "Ελεγχος ΤΕΛΙΚΟΣ ΛΑΘΟΣ" + Chr(13) + "EGGTIM"
'   GDB.Close
'
'   Exit Sub
'End If
'
'r.Close
'
'








End Sub

Private Sub Form_Load()

 
'Open "c:\mercpath.txt" For Input As #1
'Input #1, fRServer
'Input #1, gconnect
'Close #1
 
 
 'GREM.Open fConnString
 
 
 
' oPackage.FailOnError = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
'
End Sub


Private Sub Timer1_Timer()
'Th


End Sub
