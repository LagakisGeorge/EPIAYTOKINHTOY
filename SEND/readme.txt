

������� ��� c:\replica.txt  ��� remote server

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



�� ��������� ��� odbc  ��� GDB  :
GDB.Open "DSN=FORITO;UID=sa;PWD=sa"


connection string :
Provider=SQLOLEDB.1;Password=12345678;Persist Security Info=True;User ID=sa;Initial Catalog=mercury;Data Source=LENOVO-PC\sqlexpress,50172

  .SourceDatabase = "MERCURY"
  .DestinationDatabase = "MERCFORHTOY"


�������� ���� ������� ���� ���� "MERCFORHTOY" ���  remote


Package :

  .AddObjectForTransfer "BARCODES", "dbo", DTSSQLObj_UserTable
  .AddObjectForTransfer "EID", "dbo", DTSSQLObj_UserTable
  .AddObjectForTransfer "EGG", "dbo", DTSSQLObj_UserTable
  .AddObjectForTransfer "PEL", "dbo", DTSSQLObj_UserTable
  .AddObjectForTransfer "EGGTIM", "dbo", DTSSQLObj_UserTable
  .AddObjectForTransfer "TIM", "dbo", DTSSQLObj_UserTable


������ ������� ��� ����� � �����������
'================================================================
���� �� ������������� ��� �� ��������� ��� ������ ��� �� ����� ���������
������ ���������
pel,eid,egg,BARCODES  ���� ��� ��������
eggtim,tim ���� ��� �������� ��� ��������������

'=========================================================================================


��������� � ������ :



c = InputBox("���� ������ ��������������")

c2 = InputBox("���� �� �������� ��� ��������� ��� �������������� ��� ������")

c3 = "�" + Format(c, "000000")

GDB.Open "DSN=FORITO;UID=sa;PWD=sa"


r.Open "select sum(POSO) from MERCFORHTOY.dbo.EGGTIM WHERE LEFT(ATIM,7)='" + c3 + "'", GDB, adOpenDynamic, adLockOptimistic

If r(0) = Val(c2) Then
   List2.AddItem "������� �1  ok" + Chr(13) + "MERCFORHTOY.EGGTIM"
Else
   MsgBox "������� �1 �����" + Chr(13) + "MERCFORHTOY.EGGTIM"
   GDB.Close

   Exit Sub
End If
GDB.Execute "DROP TABLE PEL"
GDB.Execute "SELECT *  INTO PEL  FROM MERCFORHTOY.dbo.PEL", n


If n > 0 Then
 List2.AddItem "������� �2 ��" + Chr(13) + "PEL �������� " + Str(n)
Else
   MsgBox "������� �2 �����" + Chr(13) + "PEL "
   GDB.Close

   Exit Sub
End If



n = 0



GDB.Execute "DROP TABLE EID"
 GDB.Execute "SELECT *  INTO EID  FROM MERCFORHTOY.dbo.EID", n


If n > 0 Then
    List2.AddItem "������� �3 ok" + Chr(13) + "���� �������� " + Str(n)
Else
   MsgBox "������� �3 �����" + Chr(13) + "���� "
   GDB.Close

   Exit Sub
End If




n = 0



GDB.Execute "DROP TABLE BARCODES"
GDB.Execute "SELECT *  INTO BARCODES  FROM MERCFORHTOY.dbo.BARCODES", n

If n > 0 Then
    List2.AddItem "������� �4 ok" + Chr(13) + "BARCODES �������� " + Str(n)
Else
   MsgBox "������� �4 �����" + Chr(13) + "BARCODES "
   GDB.Close

   Exit Sub
End If




n = 0




GDB.Execute "DROP TABLE EGGTIM"
GDB.Execute "SELECT *  INTO EGGTIM   FROM MERCFORHTOY.dbo.EGGTIM WHERE ATIM='" + c3 + "'", n

If n > 0 Then
    List2.AddItem "������� �5 ok" + Chr(13) + "EGGTIM �������� " + Str(n)
Else
   MsgBox "������� �5 �����" + Chr(13) + "EGGTIM "
   GDB.Close

   Exit Sub
End If




n = 0



GDB.Execute "DROP TABLE TIM"
GDB.Execute "SELECT *  INTO TIM   FROM MERCFORHTOY.dbo.TIM WHERE ATIM='" + c3 + "'", n


If n > 0 Then
    List2.AddItem "������� �6 ok" + Chr(13) + "TIM �������� " + Str(n)
Else
   MsgBox "������� �6 �����" + Chr(13) + "TIM "
   GDB.Close

   Exit Sub
End If




n = 0





GDB.Execute "DROP TABLE EGG"
GDB.Execute "SELECT *  INTO EGG   FROM MERCFORHTOY.dbo.EGG", n

If n > 0 Then
    List2.AddItem "������� �7 ok" + Chr(13) + "EGG �������� " + Str(n)
Else
   MsgBox "������� �7 �����" + Chr(13) + "EGG "
   GDB.Close

   Exit Sub
End If



n = 0





GDB.Execute "UPDATE EGG SET IDTIM=999", n







r.Open "select sum(POSO) from EGGTIM WHERE LEFT(ATIM,7)='" + c3 + "'", GDB, adOpenDynamic, adLockOptimistic

If r(0) = Val(c2) Then
   List2.AddItem "������� �������   ok EGGTIM"
Else
   MsgBox "������� ������� �����" + Chr(13) + "EGGTIM"
   GDB.Close

   Exit Sub
End If

r.Close












