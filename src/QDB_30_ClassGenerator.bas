Attribute VB_Name = "QDB_30_ClassGenerator"
Option Explicit

' ============================================================
' Quick DB Engine
' QDB_30_ClassGenerator
'
' Class Module生成を担当する。
'
' ワークシート上の実際の位置は意識しない。
' フィールド名・型の取得はQDB_10_Schemaに任せる。
'
' そのため、
'
'   QDB_TABLE_COL
'
' が変更されても、このGenerator自身の修正は不要。
'
' ============================================================


' ============================================================
' Class Module生成
' ============================================================

Public Sub QDB_GenerateClass( _
    ByVal ws As Worksheet, _
    ByVal ClassName As String, _
    ByVal fieldCount As Long)

    Dim classModule As Object
    Dim codeStr As String

    Debug.Print _
        "  ClassGenerator : " & ClassName

    Set classModule = _
        QDB_GetOrCreateClassModule(ClassName)

    If classModule Is Nothing Then

        MsgBox _
            "Class Module [" & _
            ClassName & _
            "] を作成できませんでした。", _
            vbExclamation

        Exit Sub

    End If

    codeStr = _
        QDB_BuildClassCode( _
            ws, _
            ClassName, _
            fieldCount)

    With classModule.CodeModule

        If .CountOfLines > 0 Then

            .DeleteLines _
                1, _
                .CountOfLines

        End If

        .AddFromString codeStr

    End With

    Debug.Print _
        "  Class Module generated : " & _
        ClassName

End Sub


' ============================================================
' Class Module取得 / 作成
' ============================================================

Private Function QDB_GetOrCreateClassModule( _
    ByVal ClassName As String) As Object

    Dim classModule As Object

    On Error Resume Next

    Set classModule = _
        ThisWorkbook.VBProject.VBComponents( _
            ClassName)

    On Error GoTo 0

    If classModule Is Nothing Then

        On Error GoTo CreateError

        Set classModule = _
            ThisWorkbook.VBProject.VBComponents.Add(2)

        classModule.Name = ClassName

        On Error GoTo 0

    End If

    Set QDB_GetOrCreateClassModule = _
        classModule

    Exit Function


CreateError:

    Debug.Print _
        "  Class Module creation error : " & _
        Err.Number & _
        " / " & _
        Err.Description

    Set QDB_GetOrCreateClassModule = _
        Nothing

End Function


' ============================================================
' Class Moduleコード生成
' ============================================================

Private Function QDB_BuildClassCode( _
    ByVal ws As Worksheet, _
    ByVal ClassName As String, _
    ByVal fieldCount As Long) As String

    Dim codeStr As String

    Dim i As Long

    Dim FieldName As String
    Dim FieldType As String

    codeStr = _
        "Option Explicit" & _
        vbCrLf & vbCrLf


    ' --------------------------------------------------------
    ' フィールド定義
    ' --------------------------------------------------------

    For i = 1 To fieldCount

        FieldName = _
            QDB_GetFieldName(ws, i)

        FieldType = _
            QDB_GetFieldType(ws, i)

        codeStr = _
            codeStr & _
            "Public " & _
            FieldName & _
            " As " & _
            FieldType & _
            vbCrLf

    Next i


    codeStr = _
        codeStr & _
        vbCrLf


    ' --------------------------------------------------------
    ' 初期化Sub
    ' --------------------------------------------------------

    codeStr = _
        codeStr & _
        "Public Sub init_" & _
        ClassName & "("


    For i = 1 To fieldCount

        FieldName = _
            QDB_GetFieldName(ws, i)

        FieldType = _
            QDB_GetFieldType(ws, i)

        codeStr = _
            codeStr & _
            FieldName & _
            " As " & _
            FieldType

        If i < fieldCount Then

            codeStr = _
                codeStr & ", "

        End If

    Next i


    codeStr = _
        codeStr & _
        ")" & _
        vbCrLf


    ' --------------------------------------------------------
    ' フィールドへの代入
    ' --------------------------------------------------------

    For i = 1 To fieldCount

        FieldName = _
            QDB_GetFieldName(ws, i)

        codeStr = _
            codeStr & _
            "    Me." & _
            FieldName & _
            " = " & _
            FieldName & _
            vbCrLf

    Next i


    codeStr = _
        codeStr & _
        "End Sub" & _
        vbCrLf


    QDB_BuildClassCode = _
        codeStr


    Debug.Print _
        "  ----- Generated Class -----"

    Debug.Print _
        codeStr

    Debug.Print _
        "  ----------------------------"

End Function

