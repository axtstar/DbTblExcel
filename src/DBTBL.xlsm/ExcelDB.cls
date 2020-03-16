VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public cnt As Long
Public target As AnalizeExcel

Const adParamInput = 1
Const adInteger = 3
Const adDate = 7
Const adBSTR = 8
Const adBoolean = 11
Const adBigInt = 20
Const adDBTimeStamp = 135

'文字列化
Function getString(tp As String, str As Variant) As String
   If InStr(tp, "timestamp") > 0 Then
    'dateに一旦格納
    Dim dt As Date
    dt = CDate(str)
    getString = Format(dt, "yyyy-mm-dd hh:nn:ss")
    Else
    getString = str
   End If

End Function

Function TableCount(cSheet As Worksheet) As String

    Dim db, rec As Object

    Set target = New AnalizeExcel
    Call target.Init(cSheet)

    'データベースの初期化
    Set db = CreateObject("DB.connection")
    db.ConnectionString = target.ConnectionString
    db.Open
    
    'クエリーの実行前に件数を取得
    Set rec = db.Execute("select count(*) AS cnt from " & target.tableName)
    cnt = rec.Fields("cnt")
    
    If cnt = 0 Then
        TableCount = "データが無い"
        Exit Function
    End If

    'クエリーの結果の破棄
    rec.Close    '必要ない時もあり
    Set rec = Nothing   'こっちは常に必要
    
    'データベースの終了処理
    db.Close
    Set db = Nothing

    
End Function
 

Function Export(cSheet As Worksheet) As String

    Dim db, rec As Object

    Set target = New AnalizeExcel
    Call target.Init(cSheet)
    
    Dim max As Long
    
    If cSheet.Cells(12, 2).Value <> "" Then
        max = cSheet.Cells(12, 2).Value
    End If

    '条件を追加
    Dim s As String
    Dim i As Integer
    
    Dim j As String
    Dim c As Variant
    
    On Error Resume Next
    
    j = ""
    If target.Condition.Count > 0 Then
        j = " Where "
        For i = 1 To target.Columns.Count
            If target.Condition.item(target.Columns(i)) <> "" Then
                j = j & target.Columns(i) & "= '" & target.Condition.item(target.Columns(i)) & "' AND "
            End If
        Next
        If Len(j) > 4 Then
            j = Mid(j, 1, Len(j) - 5)
        End If
    ElseIf cSheet.Cells(13, 2) <> "" Then
        j = " Where "
    End If
    
    If target.GetOrderString() <> "" Then
        j = j & " Order by " & target.GetOrderString()
    End If

    'データベースの初期化
    Set db = CreateObject("ADODB.connection")
    db.ConnectionString = target.ConnectionString
    db.Open
    
    Dim b13 As String
    If cSheet.Cells(13, 2) <> "" Then
        b13 = cSheet.Cells(13, 2)
    End If
    
    'クエリーの実行前に件数を取得
    Set rec = db.Execute("select count(*) AS cnt from " & target.tableName & j & b13)
    cnt = rec.Fields("cnt")
    
    If cnt = 0 Then
        Export = "データが無い"
        Exit Function
    End If
    
    Set rec = db.Execute("select * from " & target.tableName & j & b13)
    
    '結果の出力
    Dim startRow, startLocate As Long
    startRow = 21
    startLocate = startRow
    Dim num, realNum As Long
    num = 1
    realNum = num
    
    '高速化のため、配列に一旦いれる(ただし上限は10000とする)
    
    Dim arraySize, arraySizeMinusOne As Integer
    arraySize = 10000
    arraySizeMinusOne = arraySize - 1
    
    Dim tempWork() As Variant
    ReDim tempWork(arraySizeMinusOne, target.ColSize)
   
    Do While Not rec.EOF
        tempWork(num - 1, 0) = realNum
        For i = 1 To target.ColSize
            If target.IsDB2 = True And target.Types(target.Columns(i)) = "TIME" Then
                tempWork(num - 1, i) = TimeFromTime(rec.Fields(target.Columns(i).Value))
            ElseIf target.Types(target.Columns(i)) = "binary" Then
                tempWork(num - 1, i) = bin2str(rec.Fields(target.Columns(i).Value), target.Columns(i).Value)
            Else
                tempWork(num - 1, i) = rec.Fields(target.Columns(i).Value)
            End If
        Next
        startRow = startRow + 1
        num = num + 1
        realNum = realNum + 1
        
        If realNum > max Then
            Exit Do
        End If
        
        '10000で割り切れたときは出力する
        If num Mod arraySize = 0 Then
            'RangeCopy
            target.currentSheet.Range(target.currentSheet.Cells(startLocate, 3), _
            target.currentSheet.Cells(startLocate + arraySizeMinusOne, 3 + target.ColSize)) = tempWork
            num = 1
            startLocate = startLocate + arraySizeMinusOne
            ReDim tempWork(arraySizeMinusOne, target.ColSize)
            Application.StatusBar = realNum & "行目を処理をしています..."
        End If
        
        rec.MoveNext
    Loop
    
    'RangeCopy
    target.currentSheet.Range(target.currentSheet.Cells(startLocate, 3), target.currentSheet.Cells(startRow - 1, 3 + target.ColSize)) = tempWork
    
    'クエリーの結果の破棄
    rec.Close    '必要ない時もあり
    Set rec = Nothing   'こっちは常に必要
    
    'データベースの終了処理
    db.Close
    Set db = Nothing


End Function

Function bin2str(bin As Variant, column As String)
    Dim ret As String
    Dim i As Long
    
    Dim b As Variant
    
    b = bin.Value
    
    For i = LBound(b) To UBound(b)
        ret = ret & Right("0" & Hex(b(i)), 2)
    Next
    
    ret = "0x" & ret
    
    bin2str = ret
End Function

'''''''''''''''''''''''''''''''''''
'IMPORTバッチ作成処理
'
Function CreateBat(cSheet As Worksheet) As String
    
    Set target = New AnalizeExcel
    Call target.Init(cSheet)
    
    Dim realNum As Long
    Dim num As Long
    num = 1
    realNum = num
    
    
    '列数と行数の取得
    cnt = target.GetRowSize()
    If Not (cSheet.OLEObjects("CheckBox1").Object.Value) And cnt = 0 Then
        CreateBat = "データが無い"
        Exit Function
    End If
    
    Dim tmpWork() As Variant
    
    '結果の出力
    Dim startRow, startLocate As Long
    startRow = 21
    startLocate = startRow
    
    '高速化のため、配列に一旦いれる(ただし上限は10000とする)
    Dim arraySize, arraySizeMinusOne As Integer
    arraySize = 5
    arraySizeMinusOne = arraySize - 1

    ReDim tempWork(arraySizeMinusOne, target.ColSize)
    tmpWork = target.currentSheet.Range(target.currentSheet.Cells(startLocate, 4), target.currentSheet.Cells(startLocate + arraySizeMinusOne, 3 + target.ColSize))

    'CSVファイルに出力
    'とりあえずカレントに
    
    
   Dim f As New FileUtility
   f.Init
   '上書きで開きます
   Call f.OpenFile(ActiveWorkbook.Path & "\" & target.tableName & ".csv", 2)
    
   'sqlを出力
   Dim f2 As New FileUtility
   f2.Init
   '上書きで開きます
   Call f2.OpenFile(ActiveWorkbook.Path & "\" & target.tableName & ".sql", 2)
   
   If cSheet.OLEObjects("CheckBox1").Object.Value = True Then
        f2.WriteLine ("delete from " & target.tableName & ";")
   End If
    
    
    '列側
    Dim i As Long
    Dim j As Long
    i = 1
    j = 1
    num = 1
    For j = 1 To cnt
        For i = 1 To target.ColSize
           'CSV出力
           If IsNumeric(target.Numbers(i).Value2) Then '数値でない場合は出力スキップ
                If target.Types(i) = "bigint" Or target.Types(i) = "integer" Or target.Types(i) = "double precision" Then
                    f.WriteEx getString(target.Types(i), tmpWork(num, i))
                Else
                    f.WriteEx """" & Replace(getString(target.Types(i), tmpWork(num, i)), """", """""") & """"
                End If

                If i <> target.ColSize Then
                    f.WriteEx target.Delimiter
                End If
           End If
        Next
        
        'SQL出力
        Dim delValue As String
        Dim rightValue As String
        
        Dim deleteSql As String
        Dim insertSql As String
        
        rightValue = JoinCnnmaEx(tmpWork, num, target)
        delValue = JoinDelEx(tmpWork, num, target)
        
        If cSheet.OLEObjects("CheckBox1").Object.Value <> True Then
            deleteSql = "delete from " & target.tableName & " where " & delValue & ";"
            f2.WriteLine (deleteSql)
        End If
        
        insertSql = "insert into " & target.tableName & " values (" & rightValue & ") ;"
        f2.WriteLine (insertSql)
        f.WriteEx vbCrLf
        If j Mod arraySize = 0 Then
            startLocate = startLocate + arraySize
            tmpWork = target.currentSheet.Range(target.currentSheet.Cells(startLocate, 4), _
            target.currentSheet.Cells(startLocate + arraySizeMinusOne, 3 + target.ColSize))
            ReDim tempWork(arraySizeMinusOne, target.ColSize)
            num = 1
        Else
            num = num + 1
        End If
    Next

    f.CloseEx
    f2.CloseEx
    Set f2 = Nothing

    'WSFの出力
    'sqlを出力
    Dim f3 As New FileUtility
    f3.Init
    '上書きで開きます
    Call f3.OpenFile(ActiveWorkbook.Path & "\" & target.tableName & ".wsf", 2)
    
    Dim writeVbStr As String
    
    writeVbStr = Replace(target.VbsString, "@rDB", target.ConnectionString)
    writeVbStr = Replace(writeVbStr, "@FilePath", target.tableName & ".sql")
    writeVbStr = Replace(writeVbStr, vbLf, vbCrLf)
    
    f3.WriteEx (writeVbStr)
    
    f3.CloseEx

    '出力用SQL生成
    Dim f4 As New FileUtility
    f4.Init
    '上書きで開きます
    Call f4.OpenFile(ActiveWorkbook.Path & "\" & target.tableName & ".pgpsql", 2)
    
    Dim writeInsertSql As String
    
    writeInsertSql = Replace(target.InsertString, "@TABLE", target.tableName)
    writeInsertSql = Replace(writeInsertSql, "@WHERE", target.WhareString)
    writeInsertSql = Replace(writeInsertSql, "@COLUMNS", target.ColumnString)
    writeInsertSql = Replace(writeInsertSql, vbLf, vbCrLf)
    
    f4.WriteEx (writeInsertSql)
    
    f4.CloseEx


End Function

Function UpdateDB2(cSheet As Worksheet)

    Set target = New AnalizeExcel
    Call target.Init(cSheet)
    
    Dim r As Long
    
    '列数と行数の取得
    cnt = target.GetRowSize()
    If cnt = 0 Then
        UpdateDB2 = "データが無い"
        Exit Function
    End If
    
    Dim tmpWork() As Variant
    
    tmpWork = target.currentSheet.Range(target.currentSheet.Cells(21, 4), target.currentSheet.Cells(20 + cnt, 3 + target.ColSize))

    Dim db, rec, command As Object

    'データベースの初期化
    Set db = CreateObject("ADODB.Connection")
    db.ConnectionString = target.ConnectionString
    db.Open
    
    Dim i As Long
    
    If target.currentSheet.OLEObjects("CheckBox1").Object.Value = True Then
        Set rec = db.Execute("delete from " & target.tableName)
    End If
        
    For i = 1 To cnt
        Set command = CreateObject("ADODB.Command")
        command.ActiveConnection = db

        Dim ParamValue As String
        Dim delValue As String
        delValue = JoinDelEx(tmpWork, i, target)
        
        Set rec = db.Execute("delete from " & target.tableName & " where " & delValue)

        
        ParamValue = ParamValues(target)
        Call JoinCnnmaEx2(tmpWork, i, target, command)
        
        command.CommandText = "insert into " & target.tableName & " (" & JoinCoulmnName(target) & ") values (" & ParamValue & ") "
        
        command.Execute
    Next
    
    'クエリーの結果の破棄
    'rec.Close    '必要ない時もあり
    Set rec = Nothing   'こっちは常に必要
    Set command = Nothing
    
    'データベースの終了処理
    db.Close
    Set db = Nothing

End Function


Function UpdateDB(cSheet As Worksheet)

    Set target = New AnalizeExcel
    Call target.Init(cSheet)
    
    Dim r As Long
    
    '列数と行数の取得
    cnt = target.GetRowSize()
    If cnt = 0 Then
        UpdateDB = "データが無い"
        Exit Function
    End If
    
    Dim tmpWork() As Variant
    
    tmpWork = target.currentSheet.Range(target.currentSheet.Cells(21, 4), target.currentSheet.Cells(20 + cnt, 3 + target.ColSize))

    Dim db, rec As Object

    'データベースの初期化
    Set db = CreateObject("ADODB.connection")
    db.ConnectionString = target.ConnectionString
    db.Open
        
    Dim i As Long
    
    If target.currentSheet.OLEObjects("CheckBox1").Object.Value = True Then
        Set rec = db.Execute("delete from " & target.tableName)
    End If
        
    For i = 1 To cnt
        Dim rightValue As String
        Dim delValue As String
        
        rightValue = JoinCnnmaEx(tmpWork, i, target)
        delValue = JoinDelEx(tmpWork, i, target)
        
        Set rec = db.Execute("delete from " & target.tableName & " where " & delValue)
        Set rec = db.Execute("insert into " & target.tableName & "(" & JoinCoulmnName(target) & ") values (" & rightValue & ") ")

    Next
    
    'クエリーの結果の破棄
    'rec.Close    '必要ない時もあり
    Set rec = Nothing   'こっちは常に必要
    
    'データベースの終了処理
    db.Close
    Set db = Nothing

End Function

Function isMemberOfCollection(col As Collection, query) As Boolean
    Dim item
    
    For Each item In col
        If item = query Then
            isMemberOfCollection = True
            Exit Function
        End If
    Next
    
    isMemberOfCollection = False
    
End Function

Private Function ParamValues(aExcel As AnalizeExcel) As String
    Dim i As Integer
    
    For i = 1 To aExcel.ColSize
    If i = aExcel.ColSize Then
        ParamValues = ParamValues + " ?"
    Else
        ParamValues = ParamValues + " ?,"
    End If
    Next

End Function

Private Function JoinCoulmnName(aExcel As AnalizeExcel)
    Dim i As Integer
    
    For i = 1 To aExcel.ColSize
        JoinCoulmnName = JoinCoulmnName & aExcel.Columns(i)
        If i <> aExcel.ColSize Then
            JoinCoulmnName = JoinCoulmnName & ","
        End If
    Next
End Function

Private Function JoinCnnmaEx(ta() As Variant, r As Long, aExcel As AnalizeExcel)
    Dim i As Integer
    
    For i = 1 To aExcel.ColSize
        If ta(r, i) = "" And isMemberOfCollection(aExcel.NotNull, aExcel.Columns(i)) = False Then
            JoinCnnmaEx = JoinCnnmaEx & "null"
        ElseIf aExcel.IsDB2 = True And aExcel.Types(aExcel.Columns(i)) = "BIGINT" Then
            JoinCnnmaEx = JoinCnnmaEx & ta(r, i)
        ElseIf aExcel.IsDB2 = True And aExcel.Types(aExcel.Columns(i)) = "SMALLINT" Then
            JoinCnnmaEx = JoinCnnmaEx & ta(r, i)
        ElseIf aExcel.IsDB2 = True And aExcel.Types(aExcel.Columns(i)) = "INTEGER" Then
            JoinCnnmaEx = JoinCnnmaEx & ta(r, i)
        ElseIf aExcel.IsDB2 = True And aExcel.Types(aExcel.Columns(i)) = "DECIMAL" Then
            JoinCnnmaEx = JoinCnnmaEx & ta(r, i)
        ElseIf aExcel.IsDB2 = True And aExcel.Types(aExcel.Columns(i)) = "TIME" Then
            JoinCnnmaEx = JoinCnnmaEx & "'" & TimeFromTime(ta(r, i)) & "'"
        ElseIf aExcel.Types(aExcel.Columns(i)) = "binary" Then
            JoinCnnmaEx = JoinCnnmaEx & TimeFromTime(ta(r, i))
        Else
            JoinCnnmaEx = JoinCnnmaEx & "'" & ta(r, i) & "'"
        End If
        If i <> aExcel.ColSize Then
            JoinCnnmaEx = JoinCnnmaEx & ","
        End If
    Next
    
End Function

Private Function JoinCnnmaEx2(ta() As Variant, r As Long, aExcel As AnalizeExcel, ByRef cmd1 As Variant)
    Dim i As Integer
    
    For i = 1 To aExcel.ColSize
        Dim param As Object
        If aExcel.Types(aExcel.Columns(i)) = "bigint" Then
            Set param = cmd1.CreateParameter(aExcel.Columns(i), adBigInt, adParamInput, , ta(r, i))
            cmd1.Parameters.Append param
        ElseIf aExcel.Types(aExcel.Columns(i)) = "integer" Then
            ' addInteger=3
            Set param = cmd1.CreateParameter(aExcel.Columns(i), adInteger, adParamInput, , ta(r, i))
            cmd1.Parameters.Append param
        ElseIf aExcel.Types(aExcel.Columns(i)) = "text" Or aExcel.Types(aExcel.Columns(i)) = "character varying" Then
            Set param = cmd1.CreateParameter(aExcel.Columns(i), adBSTR, adParamInput, , ta(r, i))
            cmd1.Parameters.Append param
        ElseIf aExcel.Types(aExcel.Columns(i)) = "boolean" Then
            Set param = cmd1.CreateParameter(aExcel.Columns(i), adBoolean, adParamInput, , ta(r, i))
            cmd1.Parameters.Append param
        ElseIf aExcel.Types(aExcel.Columns(i)) = "date" Then
            Set param = cmd1.CreateParameter(aExcel.Columns(i), adDate, adParamInput, , ta(r, i))
            cmd1.Parameters.Append param
        ElseIf aExcel.Types(aExcel.Columns(i)) = "timestamp without time zone" Then
            Set param = cmd1.CreateParameter(aExcel.Columns(i), adDBTimeStamp, adParamInput, , ta(r, i))
            cmd1.Parameters.Append param
        Else
            Set param = cmd1.CreateParameter(aExcel.Columns(i), 3, 1, , ta(r, i))
            cmd1.Parameters.Append param
        End If
    Next
End Function

Private Function TimeFromTime(var As Variant) As String
    Dim re As New RegExp
    Dim mc As MatchCollection
    Dim str As String
    re.Pattern = "([0-9][0-9]:[0-9][0-9]:[0-9][0-9]|[0-9]:[0-9][0-9]:[0-9][0-9])(AM|PM)*"
    re.Global = False
    
    If IsNull(var) Then
        TimeFromTime = ""
        Exit Function
    End If
    
    Dim now As Date
    now = DateTime.Date
    
    If var = now Then
        TimeFromTime = "00:00:00"
        Exit Function
    End If
    
    str = CStr(var)
    
    Set mc = re.Execute(str)
    If mc.Count = 1 Then
        TimeFromTime = mc(0).Value
    ElseIf mc.Count = 2 Then
        TimeFromTime = mc(0).Value
    Else
        TimeFromTime = var
    End If
End Function

Private Function JoinEx(ta() As Variant, Delimiter As String, r As Long, columnRow As Integer)
    Dim i As Integer
    
    For i = 1 To columnRow
        JoinEx = JoinEx & "'" & ta(r, i) & "'"
        If i <> columnRow Then
            JoinEx = JoinEx & Delimiter
        End If
    Next
    
End Function

Private Function JoinDelEx(ta() As Variant, r As Long, aNal As AnalizeExcel)
    Dim i As Integer
    
    For i = 1 To aNal.PrimaryKeys.Count
        If aNal.IsDB2 = True And aNal.Types(aNal.PrimaryKeys(i)) = "BIGINT" Then
            JoinDelEx = JoinDelEx & aNal.PrimaryKeys(i) & " = " & ta(r, aNal.RevertColumns(aNal.PrimaryKeys(i)))
        ElseIf aNal.IsDB2 = True And aNal.Types(aNal.PrimaryKeys(i)) = "SMALLINT" Then
            JoinDelEx = JoinDelEx & aNal.PrimaryKeys(i) & " = " & ta(r, aNal.RevertColumns(aNal.PrimaryKeys(i)))
        ElseIf aNal.IsDB2 = True And aNal.Types(aNal.PrimaryKeys(i)) = "INTEGER" Then
            JoinDelEx = JoinDelEx & aNal.PrimaryKeys(i) & " = " & ta(r, aNal.RevertColumns(aNal.PrimaryKeys(i)))
        ElseIf aNal.IsDB2 = True And aNal.Types(aNal.PrimaryKeys(i)) = "DECIMAL" Then
            JoinDelEx = JoinDelEx & aNal.PrimaryKeys(i) & " = " & ta(r, aNal.RevertColumns(aNal.PrimaryKeys(i)))
        ElseIf aNal.IsDB2 = True And aNal.Types(aNal.PrimaryKeys(i)) = "TIME" Then
            JoinDelEx = JoinDelEx & aNal.PrimaryKeys(i) & " = '" & TimeFromTime(ta(r, aNal.RevertColumns(aNal.PrimaryKeys(i)))) & "'"
        Else
            JoinDelEx = JoinDelEx & aNal.PrimaryKeys(i) & " = '" & ta(r, aNal.RevertColumns(aNal.PrimaryKeys(i))) & "'"
        End If
        
        If i <> aNal.PrimaryKeys.Count Then
            JoinDelEx = JoinDelEx & " and "
        End If
    Next
    
End Function

