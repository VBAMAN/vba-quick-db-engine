Attribute VB_Name = "QDB_50_Finder"
Option Explicit

' ============================================================
' Quick DB Engine
' QDB_50_Finder
'
' DB検索を担当する固定Engine。
'
' このモジュールはGeneratorによって生成しない。
'
' 検索対象のCollectionを受け取り、
' フィールド名と値を指定して検索する。
'
' 対応：
'   ・数値検索
'       =
'       <>
'       >
'       <
'       >=
'       <=
'
'   ・文字列検索
'       =
'       <>
'
' ============================================================


' ============================================================
' 数値検索
' ============================================================
Public Function QDB_FindRecords_Number( _
    ByVal data As Collection, _
    ByVal Field As String, _
    ByVal OperatorType As String, _
    ByVal Value As Double) As Collection

    Dim results As New Collection
    Dim item As Object
    Dim i As Long
    Dim fieldValue As Variant

    If data Is Nothing Then
        Set QDB_FindRecords_Number = results
        Exit Function
    End If

    For i = 1 To data.Count

        Set item = data.item(i)

        fieldValue = _
            CallByName(item, Field, VbGet)

        If QDB_CompareValues( _
                fieldValue, _
                Value, _
                OperatorType, _
                True) Then

            results.Add item

        End If

    Next i

    Set QDB_FindRecords_Number = results

End Function


' ============================================================
' 文字列検索
' ============================================================
Public Function QDB_FindRecords_String( _
    ByVal data As Collection, _
    ByVal Field As String, _
    ByVal Value As String) As Collection

    Dim results As New Collection
    Dim item As Object
    Dim i As Long
    Dim fieldValue As Variant

    If data Is Nothing Then
        Set QDB_FindRecords_String = results
        Exit Function
    End If

    For i = 1 To data.Count

        Set item = data.item(i)

        fieldValue = _
            CallByName(item, Field, VbGet)

        If QDB_CompareValues( _
                fieldValue, _
                Value, _
                "=", _
                False) Then

            results.Add item

        End If

    Next i

    Set QDB_FindRecords_String = results

End Function


' ============================================================
' 比較処理
'
' 数値：
'   =
'   <>
'   >
'   <
'   >=
'   <=
'
' 文字列：
'   =
'   <>
' ============================================================
Public Function QDB_CompareValues( _
    ByVal Value1 As Variant, _
    ByVal Value2 As Variant, _
    ByVal OperatorType As String, _
    ByVal IsNumeric As Boolean) As Boolean

    Dim tolerance As Double

    tolerance = 0.0001

    If IsNumeric Then

        Select Case OperatorType

            Case "="
                QDB_CompareValues = _
                    (Abs(Value1 - Value2) < tolerance)

            Case "<>"
                QDB_CompareValues = _
                    (Abs(Value1 - Value2) >= tolerance)

            Case ">"
                QDB_CompareValues = _
                    (Value1 > Value2)

            Case "<"
                QDB_CompareValues = _
                    (Value1 < Value2)

            Case ">="
                QDB_CompareValues = _
                    (Value1 > Value2 Or _
                     Abs(Value1 - Value2) < tolerance)

            Case "<="
                QDB_CompareValues = _
                    (Value1 < Value2 Or _
                     Abs(Value1 - Value2) < tolerance)

            Case Else

                MsgBox _
                    "無効な演算子 [" & _
                    OperatorType & _
                    "] です", _
                    vbExclamation

                QDB_CompareValues = False

        End Select

    Else

        Select Case OperatorType

            Case "="
                QDB_CompareValues = _
                    (Value1 = Value2)

            Case "<>"
                QDB_CompareValues = _
                    (Value1 <> Value2)

            Case Else

                MsgBox _
                    "文字列では [" & _
                    OperatorType & _
                    "] は使用できません", _
                    vbExclamation

                QDB_CompareValues = False

        End Select

    End If

End Function

