VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalizeExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public ConnectionString As String
Public tableName As String
Public Delimiter As String

'Public RowSize As Long
Public ColSize As Long
Public Numbers As New Collection
Public Columns As New Collection
Public Types As New Collection
Public RevertColumns As New Collection
Public PrimaryKeys As New Collection
Public NotNull As New Collection
Public Condition As New Collection
Public Orders As New Collection
Public currentSheet As Worksheet
Public VbsString As String
Public RoundIDs As Variant
Public NoNumberings As Variant
Public InsertString As String
Public WhareString As String
Public ColumnString As String
Public ReturnCode As String

Public IsDB2 As Boolean

Public LocR As Long
Public LocL As Long

Public Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False

End Function

Public Sub Init(mSheet As Worksheet)
    '対象のシート
    Set currentSheet = mSheet

    'シートから各種情報を取得
    ConnectionString = currentSheet.Cells(1, 4)
    Delimiter = currentSheet.Cells(2, 7)
    ReturnCode = currentSheet.Cells(2, 9)
    
    tableName = currentSheet.Cells(2, 4)
    
    On Error Resume Next
    '全体設定
    VbsString = Sheets("devSheet").Cells(7, 3)
    InsertString = Sheets("devSheet").Cells(9, 3)
    
    Dim connStr As String
    connStr = currentSheet.Cells(1, 4)
    WhareString = currentSheet.Cells(13, 2)
    
    If InStr(connStr, "IBMDADB2") > 0 Then
        IsDB2 = True
    Else
        IsDB2 = False
    End If

    'カラム
    Dim r As Range
    '位置3,4から右向きにCntl+Shif→動作をしたカラム値から、3を引く
    Set r = currentSheet.Range(currentSheet.Cells(3, 4), currentSheet.Cells(3, 4)).End(xlToRight)
    ColSize = r.column - 3

    Dim i As Integer
    'カラム
    For i = 1 To ColSize
        With Numbers
            .Add item:=currentSheet.Cells(3, 3 + i)
        End With
        With Columns
            .Add item:=currentSheet.Cells(5, 3 + i)
        End With
        With Types
            .Add Key:=currentSheet.Cells(5, 3 + i), item:=currentSheet.Cells(6, 3 + i)
        End With
        With RevertColumns
            .Add item:=i, Key:=currentSheet.Cells(5, 3 + i)
        End With
        If currentSheet.Cells(9, 3 + i) = "○" Then
            With PrimaryKeys
                .Add item:=currentSheet.Cells(5, 3 + i)
            End With
        End If
        If currentSheet.Cells(10, 3 + i) = "○" Then
            With NotNull
                .Add item:=currentSheet.Cells(5, 3 + i)
            End With
        End If
        If currentSheet.Cells(11, 3 + i) <> "" Then
            With Orders
                .Add item:=currentSheet.Cells(5, 3 + i), Key:=CStr(currentSheet.Cells(11, 3 + i))
            End With
        End If
        If currentSheet.Cells(13, 3 + i) <> "" Then
            With Condition
                .Add item:=currentSheet.Cells(13, 3 + i), Key:=currentSheet.Cells(5, 3 + i)
            End With
        End If
        If currentSheet.Cells(3, 3 + i) <> "○" And currentSheet.Cells(3, 3 + i) <> "" Then
            ColumnString = ColumnString + currentSheet.Cells(5, 3 + i).Value  + " ,"
        End If
    Next
    IF Len(ColumnString) > 2 Then
        ColumnString = Left(ColumnString, Len(ColumnString) - 2)
    END IF

End Sub

'現在のAnalizeExcelの内容をシートに反映する
Public Function Put2Sheet(currentSheet As Worksheet, tableName As String, ConnectionString)
    ' 全般設定
    ' ここにあるカラムは〇とマークする
    RoundIDs = Split(Sheets("devSheet").Cells(5, 3), ",")
    ' ここにあるカラムはナンバリングしない
    NoNumberings = Split(Sheets("devSheet").Cells(6, 3), ",")

    'シートに各種情報を設定
    currentSheet.Cells(1, 4) = ConnectionString
    
    currentSheet.Cells(2, 4) = tableName
    
    'カラム
    Dim i As Integer
    Dim columnName As Variant
    For Each columnName In Me.Columns
        'カラム
        i = RevertColumns(columnName)
        'No
        If IsInArray(CStr(columnName), RoundIDs) Then
            currentSheet.Cells(3, 3 + i) = "〇"
        ElseIf IsInArray(CStr(columnName), NoNumberings) Then
            currentSheet.Cells(3, 3 + i) = ""
        Else
            currentSheet.Cells(3, 3 + i) = i
        End If
        '和名
        currentSheet.Cells(4, 3 + i) = columnName
        'カラム名
        currentSheet.Cells(5, 3 + i) = columnName
        currentSheet.Cells(20, 3 + i) = columnName
        
        'データタイプ
        currentSheet.Cells(6, 3 + i) = Types(columnName)
        'プライマリキー
        If PrimaryKeys(columnName) = "○" Then
            currentSheet.Cells(9, 3 + i) = "○"
        End If
        'null許可
        If NotNull(columnName) = "○" Then
            currentSheet.Cells(10, 3 + i) = "○"
        End If
    Next


End Function


Public Function GetOrderString() As String
    Dim RetValue As String
    Dim obj As Object
    If Me.Orders.Count = 0 Then
    Else
        Dim i As Integer
        For i = 1 To Me.Orders.Count
            RetValue = RetValue & Me.Orders(CStr(i))
            If i = Me.Orders.Count Then
            Else
                RetValue = RetValue & ", "
            End If
        Next
    End If
    
    GetOrderString = RetValue
End Function

Public Function GetRowSize() As Long
    Dim t1 As Range
    Dim r As Long
    '列数と行数の取得
    
    If currentSheet.Cells(21, 3).Value = "" Then
        GetRowSize = 0
        Exit Function
    End If
    
    If currentSheet.Cells(22, 3).Value = "" Then
        GetRowSize = 1
        Exit Function
    End If
    
    Set t1 = currentSheet.Range(currentSheet.Cells(21, 3), currentSheet.Cells(21, 3)).End(xlDown)
    LocR = t1.Row
    r = LocR - 20
    GetRowSize = r
End Function
