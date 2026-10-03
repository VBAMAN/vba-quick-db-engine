Attribute VB_Name = "QDB_40_ModuleGenerator"
Option Explicit

' ============================================================
' Quick DB Engine
' QDB_40_ModuleGenerator
'
' Standard Module生成を担当する。
'
' ワークシート上の実際のテーブル位置は
' QDB_10_Schema に任せる。
'
' データ取得時も、
'
'   QDB_TABLE_COL
'
' を基準にしてフィールドを取得する。
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
' ============================================================


' ============================================================
' Standard Module生成
' ============================================================

Public Sub QDB_GenerateModule( _
    ByVal ws As Worksheet, _
    ByVal ClassName As String, _
    ByVal fieldCount As Long)

    Dim standardModule As Object
    Dim ModuleName As String
    Dim codeStr As String

    ModuleName = _
        QDB_MODULE_PREFIX & ClassName

    Debug.Print _
        "  ModuleGenerator : " & _
        ModuleName

    Set standardModule = _
        QDB_GetOrCreateStandardModule( _
            ModuleName)

    If standardModule Is Nothing Then

        MsgBox _
            "Standard Module [" & _
            ModuleName & _
            "] を作成できませんでした。", _
            vbExclamation

        Exit Sub

    End If

    codeStr = _
        QDB_BuildModuleCode( _
            ws, _
            ClassName, _
            fieldCount)

    With standardModule.CodeModule

        If .CountOfLines > 0 Then

            .DeleteLines _
                1, _
                .CountOfLines

        End If

        .AddFromString codeStr

    End With

    Debug.Print _
        "  Standard Module generated : " & _
        ModuleName

End Sub


' ============================================================
' Standard Module取得 / 作成
' ============================================================

Private Function QDB_GetOrCreateStandardModule( _
    ByVal ModuleName As String) As Object

    Dim standardModule As Object

    On Error Resume Next

    Set standardModule = _
        ThisWorkbook.VBProject.VBComponents( _
            ModuleName)

    On Error GoTo 0

    If standardModule Is Nothing Then

        On Error GoTo CreateError

        Set standardModule = _
            ThisWorkbook.VBProject.VBComponents.Add(1)

        standardModule.Name = _
            ModuleName

        On Error GoTo 0

    End If

    Set QDB_GetOrCreateStandardModule = _
        standardModule

    Exit Function


CreateError:

    Debug.Print _
        "  Standard Module creation error : " & _
        Err.Number & _
        " / " & _
        Err.Description

    Set QDB_GetOrCreateStandardModule = _
        Nothing

End Function


' ============================================================
' Standard Moduleコード生成
' ============================================================

Private Function QDB_BuildModuleCode( _
    ByVal ws As Worksheet, _
    ByVal ClassName As String, _
    ByVal fieldCount As Long) As String

    Dim codeStr As String
    Dim LastRow As Long

    LastRow = _
        QDB_GetLastDataRow( _
            ws, _
            1)

    codeStr = _
        "Option Explicit" & _
        vbCrLf & vbCrLf

    ' --------------------------------------------------------
    ' DB本体
    ' --------------------------------------------------------

    codeStr = _
        codeStr & _
        "Public result_" & _
        ClassName & _
        " As Collection" & _
        vbCrLf

    codeStr = _
        codeStr & _
        "Public " & _
        ClassName & _
        "_Field As Long" & _
        vbCrLf & vbCrLf


    ' --------------------------------------------------------
    ' CreateData
    ' --------------------------------------------------------

    codeStr = _
        codeStr & _
        QDB_BuildCreateDataCode( _
            ws, _
            ClassName, _
            fieldCount)


    codeStr = _
        codeStr & _
        vbCrLf & vbCrLf


    ' --------------------------------------------------------
    ' init
    ' --------------------------------------------------------

    codeStr = _
        codeStr & _
        QDB_BuildInitCode( _
            ws, _
            ClassName, _
            fieldCount, _
            LastRow)


    Debug.Print _
        "  ----- Generated Module -----"

    Debug.Print _
        codeStr

    Debug.Print _
        "  -----------------------------"

    QDB_BuildModuleCode = _
        codeStr

End Function


' ============================================================
' CreateDataコード生成
' ============================================================

Private Function QDB_BuildCreateDataCode( _
    ByVal ws As Worksheet, _
    ByVal ClassName As String, _
    ByVal fieldCount As Long) As String

    Dim codeStr As String

    Dim i As Long

    Dim FieldName As String
    Dim FieldType As String


    codeStr = _
        "Public Function CreateData_" & _
        ClassName & "("


    ' --------------------------------------------------------
    ' 引数
    ' --------------------------------------------------------

    For i = 1 To fieldCount

        FieldName = _
            QDB_GetFieldName( _
                ws, _
                i)

        FieldType = _
            QDB_GetFieldType( _
                ws, _
                i)

        codeStr = _
            codeStr & _
            FieldName & _
            " As " & _
            FieldType

        If i < fieldCount Then

            codeStr = _
                codeStr & _
                ", "

        End If

    Next i


    codeStr = _
        codeStr & _
        ") As " & _
        ClassName & _
        vbCrLf


    ' --------------------------------------------------------
    ' Class生成
    ' --------------------------------------------------------

    codeStr = _
        codeStr & _
        "    Dim newData As " & _
        ClassName & _
        vbCrLf

    codeStr = _
        codeStr & _
        "    Set newData = New " & _
        ClassName & _
        vbCrLf


    ' --------------------------------------------------------
    ' init呼び出し
    ' --------------------------------------------------------

    codeStr = _
        codeStr & _
        "    newData.init_" & _
        ClassName & _
        " "


    For i = 1 To fieldCount

        FieldName = _
            QDB_GetFieldName( _
                ws, _
                i)

        codeStr = _
            codeStr & _
            FieldName

        If i < fieldCount Then

            codeStr = _
                codeStr & _
                ", "

        End If

    Next i


    codeStr = _
        codeStr & _
        vbCrLf


    ' --------------------------------------------------------
    ' 戻り値
    ' --------------------------------------------------------

    codeStr = _
        codeStr & _
        "    Set CreateData_" & _
        ClassName & _
        " = newData" & _
        vbCrLf

    codeStr = _
        codeStr & _
        "End Function" & _
        vbCrLf


    QDB_BuildCreateDataCode = _
        codeStr

End Function


' ============================================================
' initコード生成
'
' ★今回の重要部分
'
' FieldIndexではなく、
' QDB_TABLE_COL + i - 1
' で実際のワークシート列を求める。
'
' ============================================================

Private Function QDB_BuildInitCode( _
    ByVal ws As Worksheet, _
    ByVal ClassName As String, _
    ByVal fieldCount As Long, _
    ByVal LastRow As Long) As String

    Dim codeStr As String

    Dim i As Long
    Dim j As Long

    Dim FieldName As String
    Dim FieldType As String
    Dim CellValue As Variant

    Dim ColumnIndex As Long


    ' --------------------------------------------------------
    ' Sub開始
    ' --------------------------------------------------------

    codeStr = _
        "Public Sub init_" & _
        ClassName & "()" & _
        vbCrLf


    codeStr = _
        codeStr & _
        "    Dim data As Collection" & _
        vbCrLf


    codeStr = _
        codeStr & _
        "    Set data = New Collection" & _
        vbCrLf & vbCrLf


    ' --------------------------------------------------------
    ' データ行
    ' --------------------------------------------------------

    For j = QDB_DATA_START_ROW To LastRow

        codeStr = _
            codeStr & _
            "    data.Add CreateData_" & _
            ClassName & _
            "("


        For i = 1 To fieldCount

            FieldName = _
                QDB_GetFieldName( _
                    ws, _
                    i)

            FieldType = _
                QDB_GetFieldType( _
                    ws, _
                    i)


            ' ------------------------------------------------
            ' ★ここが今回の修正点
            '
            ' FieldIndex 1 → QDB_TABLE_COL
            ' FieldIndex 2 → QDB_TABLE_COL + 1
            ' FieldIndex 3 → QDB_TABLE_COL + 2
            ' ------------------------------------------------

            ColumnIndex = _
                QDB_TABLE_COL + i - 1


            CellValue = _
                ws.Cells( _
                    j, _
                    ColumnIndex).Value


            ' ------------------------------------------------
            ' String
            ' ------------------------------------------------

            If StrComp( _
                    FieldType, _
                    "String", _
                    vbTextCompare) = 0 Then

                codeStr = _
                    codeStr & _
                    """" & _
                    Replace( _
                        CStr(CellValue), _
                        """", _
                        """""") & _
                    """"

            Else

                ' ------------------------------------------------
                ' 数値型
                ' ------------------------------------------------

                codeStr = _
                    codeStr & _
                    CStr(CellValue)

            End If


            If i < fieldCount Then

                codeStr = _
                    codeStr & _
                    ", "

            End If

        Next i


        codeStr = _
            codeStr & _
            ")" & _
            vbCrLf

    Next j


    codeStr = _
        codeStr & _
        vbCrLf


    ' --------------------------------------------------------
    ' Collectionを公開
    ' --------------------------------------------------------

    codeStr = _
        codeStr & _
        "    Set result_" & _
        ClassName & _
        " = data" & _
        vbCrLf


    ' --------------------------------------------------------
    ' レコード数
    ' --------------------------------------------------------

    codeStr = _
        codeStr & _
        "    " & _
        ClassName & _
        "_Field = data.Count" & _
        vbCrLf


    codeStr = _
        codeStr & _
        "End Sub" & _
        vbCrLf


    QDB_BuildInitCode = _
        codeStr

End Function


