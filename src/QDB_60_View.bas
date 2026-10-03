Attribute VB_Name = "QDB_60_View"
Option Explicit

' ============================================================
' Quick DB Engine
' QDB_60_View
'
' 検索結果の表示を担当する固定Engine。
'
' このモジュールはGeneratorによって生成しない。
'
' 検索そのものは QDB_50_Finder が担当し、
' このモジュールは取得した結果を表示する。
'
' ============================================================


' ============================================================
' 検索結果を表示
'
' Data
'   : 検索対象となるCollection
'
' Field
'   : 検索対象フィールド名
'
' Value
'   : 検索値
'
' OperatorType
'   : 検索演算子
'
' FieldNames
'   : 表示するフィールド名の配列
' ============================================================
Public Sub QDB_DataView( _
    ByVal data As Collection, _
    ByVal Field As String, _
    ByVal Value As Variant, _
    ByVal OperatorType As String, _
    ByVal FieldNames As Variant)

    Dim SearchResults As Collection
    Dim obj As Object

    Dim recordCount As Long
    Dim fieldCount As Long

    Dim resultData() As Variant

    Dim i As Long
    Dim j As Long

    Dim FieldName As String
    Dim msg As String

    ' --------------------------------------------------------
    ' データの存在確認
    ' --------------------------------------------------------
    If data Is Nothing Then

        MsgBox _
            "一致するデータが見つかりません", _
            vbExclamation

        Exit Sub

    End If

    If data.Count = 0 Then

        MsgBox _
            "一致するデータが見つかりません", _
            vbExclamation

        Exit Sub

    End If


    ' --------------------------------------------------------
    ' フィールド数の取得
    ' --------------------------------------------------------
    fieldCount = QDB_GetViewFieldCount(FieldNames)

    If fieldCount <= 0 Then

        MsgBox _
            "表示フィールドがありません。", _
            vbExclamation

        Exit Sub

    End If


    ' --------------------------------------------------------
    ' 検索結果の取得
    ' --------------------------------------------------------
    If IsNumeric( _
        CallByName(data.item(1), Field, VbGet)) Then

        Set SearchResults = _
            QDB_FindRecords_Number( _
                data, _
                Field, _
                OperatorType, _
                CDbl(Value))

    Else

        Set SearchResults = _
            QDB_FindRecords_String( _
                data, _
                Field, _
                CStr(Value))

    End If


    ' --------------------------------------------------------
    ' 検索結果の確認
    ' --------------------------------------------------------
    If SearchResults Is Nothing Then

        MsgBox _
            "一致するデータが見つかりません", _
            vbExclamation

        Exit Sub

    End If

    If SearchResults.Count = 0 Then

        MsgBox _
            "一致するデータが見つかりません", _
            vbExclamation

        Exit Sub

    End If


    ' --------------------------------------------------------
    ' 検索結果を配列へ格納
    ' --------------------------------------------------------
    recordCount = SearchResults.Count

    ReDim resultData( _
        1 To recordCount, _
        1 To fieldCount)


    For i = 1 To recordCount

        Set obj = SearchResults.item(i)

        For j = 1 To fieldCount

            FieldName = _
                CStr(FieldNames(j))

            resultData(i, j) = _
                CallByName( _
                    obj, _
                    FieldName, _
                    VbGet)

        Next j

    Next i


    ' --------------------------------------------------------
    ' 表示文字列の作成
    ' --------------------------------------------------------
    msg = "検索結果: " & vbCrLf


    For i = 1 To recordCount

        For j = 1 To fieldCount

            msg = msg & _
                CStr(FieldNames(j)) & _
                " : " & _
                CStr(resultData(i, j)) & _
                vbCrLf

        Next j

        msg = msg & vbCrLf

    Next i


    ' --------------------------------------------------------
    ' 表示
    ' --------------------------------------------------------
    MsgBox _
        msg, _
        vbInformation

End Sub


' ============================================================
' フィールド数取得
' ============================================================
Private Function QDB_GetViewFieldCount( _
    ByVal FieldNames As Variant) As Long

    Dim i As Long

    QDB_GetViewFieldCount = 0

    On Error GoTo ErrorHandler

    For i = LBound(FieldNames) To UBound(FieldNames)

        If Trim$(CStr(FieldNames(i))) <> "" Then

            QDB_GetViewFieldCount = _
                QDB_GetViewFieldCount + 1

        End If

    Next i

    Exit Function


ErrorHandler:

    QDB_GetViewFieldCount = 0

End Function

