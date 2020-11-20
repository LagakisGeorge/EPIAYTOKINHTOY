

διαβαζω απο c:\replica.txt  τον remote server

Open "c:\replica.txt" For Input As #1
Input #1, fRServer
Input #1, fRUserName
Input #1, fRPassword
Input #1, fTopikos
Input #1, fRUserName
Input #1, fRPassword
Input #1, fconnection string
Close #1




'================================================
CONFIGURATION LINKED SERVER :

LINKED SERVER : LENOVO
PROVIDER : SQL NATIVE CLIENT
PRODUCT NAME : SQL
DATASOURCE : LENOVO-PC\SQLEXPRESS,50172
PROVIDER STRING : Provider=SQLOLEDB.1;Password=12345678;Persist Security Info=True;User ID=sa;Initial Catalog=mercury;Data Source=LENOVO-PC\sqlexpress,50172

SECURITY :
BE MADE BY SECURITY   
   sa
   12345678

'================================================



θα χρειαστει και odbc  για GDB  :
GDB.Open "DSN=FORITO;UID=sa;PWD=sa"


connection string :
Provider=SQLOLEDB.1;Password=12345678;Persist Security Info=True;User ID=sa;Initial Catalog=mercury;Data Source=LENOVO-PC\sqlexpress,50172

  .SourceDatabase = "MERCURY"
  .DestinationDatabase = "MERCFORHTOY"


μεταφερω τους πίνακες στην βάση "MERCFORHTOY" του  remote


Package :

  .AddObjectForTransfer "BARCODES", "dbo", DTSSQLObj_UserTable
  .AddObjectForTransfer "EID", "dbo", DTSSQLObj_UserTable
  .AddObjectForTransfer "EGG", "dbo", DTSSQLObj_UserTable
  .AddObjectForTransfer "PEL", "dbo", DTSSQLObj_UserTable
  .AddObjectForTransfer "EGGTIM", "dbo", DTSSQLObj_UserTable
  .AddObjectForTransfer "TIM", "dbo", DTSSQLObj_UserTable


βγαζει ενδειξη οτι εγινε η επικοινωνια
'================================================================
δινω το συγκεντρωτικο και το μεταφερει στο φορητό για να πάρει ποσότητες
επίσης μεταφερει
pel,eid,egg,BARCODES  όλες τις εγγραφες
eggtim,tim μόνο τις εγγραφές του συγκεντρωτικού

'=========================================================================================


αναλυτικά η λογική :



c = InputBox("δωσε αριθμο συγκεντρωτικου")

c2 = InputBox("δωσε το αθροισμα των ποσοτήτων του συγκεντρωτικού για έλεγχο")

c3 = "τ" + Format(c, "000000")

GDB.Open "DSN=FORITO;UID=sa;PWD=sa"


r.Open "select sum(POSO) from MERCFORHTOY.dbo.EGGTIM WHERE LEFT(ATIM,7)='" + c3 + "'", GDB, adOpenDynamic, adLockOptimistic

If r(0) = Val(c2) Then
   List2.AddItem "Ελεγχος Ν1  ok" + Chr(13) + "MERCFORHTOY.EGGTIM"
Else
   MsgBox "Ελεγχος Ν1 ΛΑΘΟΣ" + Chr(13) + "MERCFORHTOY.EGGTIM"
   GDB.Close

   Exit Sub
End If
GDB.Execute "DROP TABLE PEL"
GDB.Execute "SELECT *  INTO PEL  FROM MERCFORHTOY.dbo.PEL", n


If n > 0 Then
 List2.AddItem "Ελεγχος Ν2 οκ" + Chr(13) + "PEL εγγραφες " + Str(n)
Else
   MsgBox "Ελεγχος Ν2 ΛΑΘΟΣ" + Chr(13) + "PEL "
   GDB.Close

   Exit Sub
End If



n = 0



GDB.Execute "DROP TABLE EID"
 GDB.Execute "SELECT *  INTO EID  FROM MERCFORHTOY.dbo.EID", n


If n > 0 Then
    List2.AddItem "Ελεγχος Ν3 ok" + Chr(13) + "ειδη εγγραφες " + Str(n)
Else
   MsgBox "Ελεγχος Ν3 ΛΑΘΟΣ" + Chr(13) + "ειδη "
   GDB.Close

   Exit Sub
End If




n = 0



GDB.Execute "DROP TABLE BARCODES"
GDB.Execute "SELECT *  INTO BARCODES  FROM MERCFORHTOY.dbo.BARCODES", n

If n > 0 Then
    List2.AddItem "ελεγχος Ν4 ok" + Chr(13) + "BARCODES εγγραφες " + Str(n)
Else
   MsgBox "Ελεγχος Ν4 ΛΑΘΟΣ" + Chr(13) + "BARCODES "
   GDB.Close

   Exit Sub
End If




n = 0




GDB.Execute "DROP TABLE EGGTIM"
GDB.Execute "SELECT *  INTO EGGTIM   FROM MERCFORHTOY.dbo.EGGTIM WHERE ATIM='" + c3 + "'", n

If n > 0 Then
    List2.AddItem "Ελεγχος Ν5 ok" + Chr(13) + "EGGTIM εγγραφες " + Str(n)
Else
   MsgBox "Ελεγχος Ν5 ΛΑΘΟΣ" + Chr(13) + "EGGTIM "
   GDB.Close

   Exit Sub
End If




n = 0



GDB.Execute "DROP TABLE TIM"
GDB.Execute "SELECT *  INTO TIM   FROM MERCFORHTOY.dbo.TIM WHERE ATIM='" + c3 + "'", n


If n > 0 Then
    List2.AddItem "Ελεγχος Ν6 ok" + Chr(13) + "TIM εγγραφες " + Str(n)
Else
   MsgBox "Ελεγχος Ν6 ΛΑΘΟΣ" + Chr(13) + "TIM "
   GDB.Close

   Exit Sub
End If




n = 0





GDB.Execute "DROP TABLE EGG"
GDB.Execute "SELECT *  INTO EGG   FROM MERCFORHTOY.dbo.EGG", n

If n > 0 Then
    List2.AddItem "Ελεγχος Ν7 ok" + Chr(13) + "EGG εγγραφες " + Str(n)
Else
   MsgBox "Ελεγχος Ν7 ΛΑΘΟΣ" + Chr(13) + "EGG "
   GDB.Close

   Exit Sub
End If



n = 0





GDB.Execute "UPDATE EGG SET IDTIM=999", n







r.Open "select sum(POSO) from EGGTIM WHERE LEFT(ATIM,7)='" + c3 + "'", GDB, adOpenDynamic, adLockOptimistic

If r(0) = Val(c2) Then
   List2.AddItem "Ελεγχος ΤΕΛΙΚΟΣ   ok EGGTIM"
Else
   MsgBox "Ελεγχος ΤΕΛΙΚΟΣ ΛΑΘΟΣ" + Chr(13) + "EGGTIM"
   GDB.Close

   Exit Sub
End If

r.Close












