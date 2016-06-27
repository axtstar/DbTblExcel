Attribute VB_Name = "Automation"
Option Explicit

Sub GetAllData()
    Dim exDb As New ExcelDB
    Dim message As String
    Dim sheet As Worksheet
    
    For Each sheet In Worksheets
         Application.StatusBar = sheet.Name + "��������"
        If sheet.Cells(3, 4).Value <> "" And sheet.Cells(1, 2).Value = "�ڑ����" Then
            message = exDb.Export(sheet)
        End If
    Next
    
    Set exDb = Nothing

    MsgBox ("�I��")
    Application.StatusBar = "�I��"


End Sub


Sub DeleteSheets()
    
    Dim x As Integer
    Dim bl As Boolean
    bl = True
    
    x = 2
    While bl
    If Sheets("�ݒ�").Cells(x, 1).Value = "�~" Then
        Dim shName As String
        shName = Sheets("�ݒ�").Cells(x, 2).Value
        Application.DisplayAlerts = False
        Worksheets(shName).Delete
        Sheets("�ݒ�").Rows(x).Delete
        Application.DisplayAlerts = True
        x = x - 1
    End If
    x = x + 1
    If Sheets("�ݒ�").Cells(x, 1).Value = "" Then
        bl = False
    End If
    Wend
    
End Sub

Sub GetAllTables()
    Dim db, rec As Object
    Dim cnt, total As Integer
    
    Dim limmitTable As String
    
    limmitTable = Sheets("�ݒ�").Cells(1, 6).Value

    '�f�[�^�x�[�X�̏�����
    Set db = CreateObject("ADODB.connection")
    db.ConnectionString = Sheets("�ݒ�").Cells(1, 4).Value
    db.Open
    
    '�N�G���[�̎��s�O�Ɍ������擾
    Set rec = db.Execute(GetAllTablesString(limmitTable))
    
    Dim target As New AnalizeExcel
    Dim tableName As String
    Dim shName As String
    cnt = 1
    total = 2
    Do While Not rec.EOF
        '�e�[�u�����𒲂ׂĕύX������΁A�폜����
        If tableName <> rec.Fields("TableName") Then
            '�V�[�g�ɑ΂��ď������݂��s��
            If tableName <> "" Then
                '�Ă�Ղ���R�s�[
                Sheets("�Ă�Ղ�").Copy After:=Sheets(Sheets.Count)
                Dim sht As Worksheet
                Set sht = Sheets(Sheets.Count)
                
                '�V�[�g�����e�[�u������
                shName = CStr(total - 1) + "@" + tableName
                If Len(shName) > 31 Then
                    sht.Name = Left(shName, 31)
                Else
                    sht.Name = shName
                End If
                
                '�ݒ�V�[�g�ɋL��
                Sheets("�ݒ�").Cells(total, 1).Value = total - 1
                Sheets("�ݒ�").Cells(total, 2).Value = sht.Name
                Sheets("�ݒ�").Hyperlinks.Add Anchor:=Sheets("�ݒ�").Cells(total, 2), Address:="", SubAddress:= _
                    "'" & sht.Name & "'!A1", TextToDisplay:="'" & sht.Name
                Sheets("�ݒ�").Cells(total, 5).Value = tableName
                    
                total = total + 1
                
                '�V�[�g�ɏ�������
                Call target.Put2Sheet(sht, tableName, Sheets("�ݒ�").Cells(1, 4).Value)
            End If
            
            '������
            Set target = Nothing
            Set target = New AnalizeExcel
            tableName = rec.Fields("TableName")
            cnt = 1
        End If
        
        Dim columnName As String
        
        columnName = rec.Fields("ColumnName").Value
        
        '�t�B�[���h��
        With target.Columns
            .Add item:=columnName
        End With
        '�^�C�v
        With target.Types
            .Add Key:=columnName, item:=rec.Fields("DataType").Value
        End With
        With target.RevertColumns
            .Add item:=cnt, Key:=columnName
        End With
        'PK
        If rec.Fields("PrimaryKey�\���p").Value = "��" Then
            With target.PrimaryKeys
                .Add Key:=columnName, item:="��"
            End With
        Else
            With target.PrimaryKeys
                .Add Key:=columnName, item:=""
            End With
        End If
        'Not NULL
        If rec.Fields("NotNull�\���p").Value = "��" Then
            With target.NotNull
                .Add Key:=columnName, item:="��"
            End With
        Else
            With target.NotNull
                .Add Key:=columnName, item:=""
            End With
        End If

        cnt = cnt + 1
        rec.MoveNext
    Loop
    
    '�Ō�̎c��
    If tableName <> "" Then
        '�Ă�Ղ���R�s�[
        Sheets("�Ă�Ղ�").Copy After:=Sheets(Sheets.Count)
        
        Set sht = Sheets(Sheets.Count)
        '�V�[�g�����e�[�u������
        shName = CStr(total - 1) + "@" + tableName
        If Len(shName) > 31 Then
            sht.Name = Left(shName, 31)
        Else
            sht.Name = shName
        End If
        
        '�ݒ�V�[�g�ɋL��
        Sheets("�ݒ�").Cells(total, 1).Value = total - 1
        Sheets("�ݒ�").Cells(total, 2).Value = sht.Name
        Sheets("�ݒ�").Hyperlinks.Add Anchor:=Sheets("�ݒ�").Cells(total, 2), Address:="", SubAddress:= _
            "'" & sht.Name & "'!A1", TextToDisplay:="'" & sht.Name
        Sheets("�ݒ�").Cells(total, 5).Value = tableName
        total = total + 1
        
        '�V�[�g�ɏ�������
        Call target.Put2Sheet(sht, tableName, Sheets("�ݒ�").Cells(1, 4).Value)
    End If

    
    '�N�G���[�̌��ʂ̔j��
    rec.Close    '�K�v�Ȃ���������
    Set rec = Nothing   '�������͏�ɕK�v
    
    '�f�[�^�x�[�X�̏I������
    db.Close
    Set db = Nothing
End Sub

Sub GetAllScript()
    On Error Resume Next
    
    Dim fo As Object
    Set fo = CreateObject("Scripting.FileSystemObject")
    
    Dim i As Integer
    Dim str As String
    Dim cnt As Long
    i = 2
    str = Sheets("�ݒ�").Cells(i, 2).Value
    
   'Script���o��
   Dim f As New FileUtility
   f.Init
   
   Call f.OpenFile(ActiveWorkbook.Path & "\All.wsf", 2)
    f.WriteLine ("<?xml version=""1.0"" encoding=""utf-8"" ?>")
    f.WriteLine ("<package>")
    f.WriteLine ("  <job>")
    f.WriteLine ("    <runtime></runtime>")
    f.WriteLine ("    <script language=""VBScript"">")
    f.WriteLine ("    <![CDATA[")
    
    f.WriteLine ("Dim s,fso,ts,strCurPath")
    f.WriteLine ("strCurPath = WScript.ScriptFullName")
    f.WriteLine ("Set fso=CreateObject(""Scripting.FileSystemObject"")")
    
    f.WriteLine ("Set obj = Fso.GetFile( strCurPath )")
    f.WriteLine ("Set obj = obj.ParentFolder")

    
    f.WriteLine ("IF fso.FileExists(obj.Path & ""\All.touch"") Then")
    f.WriteLine ("  fso.DeleteFile(obj.Path & ""\All.touch"")")
    f.WriteLine ("End If")
    
    f.WriteLine ("Set s=WScript.CreateObject(""WScript.Shell"")")
    
    While str <> ""
        Dim exDb As New ExcelDB
        Dim message As String
        
        If Sheets("�ݒ�").Cells(i, 4).Value = "��" Then
            message = exDb.CreateBat(Sheets(str))
            If fo.FileExists(ActiveWorkbook.Path & "\" & exDb.target.tableName & ".wsf") = True Then
                f.WriteLine ("call s.run(""wscript "" & obj.Path & ""\" & exDb.target.tableName & ".wsf"",1,1)")
            End If
            Set exDb = Nothing
        End If
        
        DoEvents
        
        i = i + 1
        str = Sheets("�ݒ�").Cells(i, 2).Value
        Application.StatusBar = Sheets("�ݒ�").Cells(i, 2).Value
    Wend
    
    f.WriteLine ("fso.CreateTextFile(obj.Path & ""\All.touch""" & ")")
    f.WriteLine ("Set ts=Nothing")
    f.WriteLine ("Set fso=Nothing")

    
    f.WriteLine ("     ]]>")
    f.WriteLine ("    </script>")
    f.WriteLine ("  </job>")
    f.WriteLine ("</package>")

    f.CloseEx
    
    
    
    Set fo = Nothing
    
End Sub


Sub GetAllCount()
    On Error Resume Next
    Dim i As Integer
    Dim str As String
    Dim cnt As Long
    i = 2
    str = Sheets("�ݒ�").Cells(i, 2).Value
    While str <> ""
        Dim exDb As New ExcelDB
        Dim message As String
        
        message = exDb.TableCount(Sheets(str))
        
        DoEvents

        '����
        Sheets("�ݒ�").Cells(i, 3).Value = exDb.cnt
        
        Set exDb = Nothing
        i = i + 1
        str = Sheets("�ݒ�").Cells(i, 2).Value
    Wend
    
    
End Sub

Sub GetSchema()

    Dim db, rec As Object

    Dim target As New AnalizeExcel
    Call target.Init(ActiveSheet)

    '�f�[�^�x�[�X�̏�����
    Set db = CreateObject("ADODB.connection")
    db.ConnectionString = target.ConnectionString
    db.Open
    
    '�N�G���[�̎��s�O�Ɍ������擾
    Dim cnt As Long
    '�t�B�[���h���������o��
    Dim i As Integer
    
    'ADOX
    Dim hCatalog As ADOX.Catalog
    Dim tbl As ADOX.Table
    Set hCatalog = New ADOX.Catalog
    hCatalog.ActiveConnection = db
    Set tbl = hCatalog.tables(target.tableName)
        
    Dim keyCol As Collection
    Dim k As ADOX.Key
    Dim ind As ADOX.Index
    
    On Error Resume Next
    
    For Each k In tbl.Keys
        If k.Name <> "" Then
            keyCol.Add (k.Name)
        End If
    Next
        
    Dim c As ADOX.Index
    For Each ind In tbl.Indexes
        If ind.PrimaryKey Then
            For i = 1 To ind.Columns.Count
                Set c = ind.Columns(i)
                keyCol.Add (c.Name)
            Next
        End If
    Next
    
    Dim col As ADOX.column
    For Each col In tbl.Columns
      Cells(21, 4 + i) = col.Name
      Cells(22, 4 + i) = col.DefinedSize
      Cells(23, 4 + i) = col.Type
      'Cells(40 + i, 7+i) = col.Properties("Primary Key").Value
      'Cells(40 + i, 8) = col.Properties("Nullable").Value
      'tbl.Indexes.Count
      
      i = i + 1
      
      
    Next
    
    

End Sub


Function TableCount()
    Dim exDb As New ExcelDB
    Dim message As String
    
    message = exDb.TableCount(ActiveSheet)

    TableCount = exDb.cnt
    
    Set exDb = Nothing

    Application.StatusBar = "�I��"

End Function


Sub CreateBat()
    Dim exDb As New ExcelDB
    Dim message As String
    
    message = exDb.CreateBat(ActiveSheet)

    If exDb.cnt = 0 Then
        MsgBox ("�f�[�^������")
    Else
        MsgBox (exDb.cnt & "��")
    End If
    
    Set exDb = Nothing

    Application.StatusBar = "�I��"

End Sub

Sub UpdateDB()
    Dim result As Integer
    
    result = MsgBox("DB�𒼐ڏ��������܂��B��낵���ł���?", vbYesNo, "Confirmation")
    
    If result = vbNo Then
        Exit Sub
    End If
    
    Dim exDb As New ExcelDB
    Dim message As String
    
    If ActiveSheet.OLEObjects("CheckBox2").Object.Value = True Then
        message = exDb.UpdateDB2(ActiveSheet)
    Else
        message = exDb.UpdateDB(ActiveSheet)
    End If

    If exDb.cnt = 0 Then
        MsgBox ("�f�[�^������")
    Else
        MsgBox (exDb.cnt & "��")
    End If
    
    Set exDb = Nothing

    Application.StatusBar = "�I��"

End Sub

'Excel�N���A
Sub ExcelClear()
    Dim currentSheet As Worksheet
    Set currentSheet = ActiveSheet
    
    Dim r, l As Long
    Dim t1, t2 As Range
    
    Set t1 = currentSheet.Range(Cells(21, 3), Cells(21, 3)).End(xlDown)
    r = t1.Row
    Set t2 = currentSheet.Range(Cells(5, 3), Cells(5, 3)).End(xlToRight)
    l = t2.column

    ActiveSheet.Range(Cells(21, 3), Cells(r, l)).Clear
    ActiveSheet.Range(Cells(21, 3), Cells(r, l)).NumberFormatLocal = "@"
End Sub

'''''''''''''''''''''''''''''''''''
'Excel�ɓǂݍ���ł��鏈��
'
Sub Export()
    Dim exDb As New ExcelDB
    Dim message As String
    
    message = exDb.Export(ActiveSheet)

    If exDb.cnt = 0 Then
        MsgBox ("�f�[�^������")
    Else
        MsgBox (exDb.cnt & "��")
    End If
    
    Set exDb = Nothing

    Application.StatusBar = "�I��"
    
End Sub


Sub SetSheetsLink()
    Dim wk As Worksheet
    Dim i As Integer
    i = 1
    For Each wk In Sheets
        Sheets("�ݒ�").Hyperlinks.Add Anchor:=Cells(i + 1, 2), Address:="", SubAddress:= _
            "'" & wk.Name & "'!A1", TextToDisplay:="'" & wk.Name & "'!A1"
        i = i + 1
    Next
End Sub

