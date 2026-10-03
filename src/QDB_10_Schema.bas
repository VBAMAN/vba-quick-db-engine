Attribute VB_Name = "QDB_10_Schema"
Option Explicit

' ============================================================
' Quick DB Engine
' QDB_10_Schema
'
' ワークシート上のDB定義位置を管理する。
'
' QDB_TABLE_COL をテーブル開始列として、
' フィールド番号をワークシート上の列番号へ変換する。
'
' 例：
'
'   QDB_TABLE_COL = 5
'
'       E       F       G
'       ID      X       Y
'       String  Long    Long
'       1000    10      20
'
'   FieldIndex 1 → E列
'   FieldIndex 2 → F列
'   FieldIndex 3 → G列
'
' Generator側は実際の列番号を意識しない。
'
' ============================================================


' ============================================================
' フィールド数取得
' ============================================================

Public Function QDB_GetFieldCount( _
    ByVal ws As Worksheet) As Long

    Dim i As Long
    Dim ColumnIndex As Long

    QDB_GetFieldCount = 0

    For i = 1 To QDB_MAX_FIELDS

        ColumnIndex = _
            QDB_TABLE_COL + i - 1

        If Trim$( _
            CStr(ws.Cells( _
                QDB_FIELD_ROW, _
                ColumnIndex).Value)) = "" Then

            Exit For

        End If

        QDB_GetFieldCount = _
            QDB_GetFieldCount + 1

    Next i

End Function


' ============================================================
' フィールド名一覧取得
' ============================================================

Public Function QDB_GetFieldNames( _
    ByVal ws As Worksheet, _
    ByVal fieldCount As Long) As Variant

    Dim FieldNames() As String
    Dim i As Long

    If fieldCount <= 0 Then

        QDB_GetFieldNames = Array()

        Exit Function

    End If

    ReDim FieldNames(1 To fieldCount)

    For i = 1 To fieldCount

        FieldNames(i) = _
            QDB_GetFieldName(ws, i)

    Next i

    QDB_GetFieldNames = FieldNames

End Function


' ============================================================
' フィールド型取得
' ============================================================

Public Function QDB_GetFieldType( _
    ByVal ws As Worksheet, _
    ByVal FieldIndex As Long) As String

    Dim ColumnIndex As Long

    ColumnIndex = _
        QDB_TABLE_COL + FieldIndex - 1

    QDB_GetFieldType = _
        Trim$( _
            CStr(ws.Cells( _
                QDB_TYPE_ROW, _
                ColumnIndex).Value))

End Function


' ============================================================
' フィールド名取得
' ============================================================

Public Function QDB_GetFieldName( _
    ByVal ws As Worksheet, _
    ByVal FieldIndex As Long) As String

    Dim ColumnIndex As Long

    ColumnIndex = _
        QDB_TABLE_COL + FieldIndex - 1

    QDB_GetFieldName = _
        Trim$( _
            CStr(ws.Cells( _
                QDB_FIELD_ROW, _
                ColumnIndex).Value))

End Function


' ============================================================
' 最終データ行取得
'
' FieldIndexはDB内のフィールド番号。
' 実際のワークシート列はQDB_TABLE_COLから算出する。
' ============================================================

Public Function QDB_GetLastDataRow( _
    ByVal ws As Worksheet, _
    ByVal FieldIndex As Long) As Long

    Dim ColumnIndex As Long

    ColumnIndex = _
        QDB_TABLE_COL + FieldIndex - 1

    QDB_GetLastDataRow = _
        ws.Cells( _
            ws.Rows.Count, _
            ColumnIndex).End(xlUp).Row

End Function

