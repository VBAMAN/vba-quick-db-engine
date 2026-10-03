Attribute VB_Name = "QDB_20_Generator"
Option Explicit

' ============================================================
' Quick DB Engine
' QDB_20_Generator
'
' Quick DB全体の生成処理を管理する。
'
' このモジュール自身はコードを生成しない。
'
' 生成するもの：
'   ・Class Module
'   ・Standard Module
'
' Finder / View はEngine固定コードとして存在するため、
' このGeneratorでは生成しない。
'
' ============================================================

Public Sub Run_Generate()

    Dim ws As Worksheet
    Dim ClassName As String

    Set ws = _
        ThisWorkbook.Worksheets( _
            QDB_SHEET_NAME)

    ClassName = ws.Name

    '------------------------------------------
    ' Class Module 名の事前チェック
    '------------------------------------------
    If Not QDB_ValidateClassName(ClassName) Then
        Exit Sub
    End If

    '------------------------------------------
    ' 既存生成物の削除
    '------------------------------------------
    QDB_PrepareGenerate ClassName

    '------------------------------------------
    ' Generator 実行
    '------------------------------------------
    QDB_Generate

End Sub


'==================================================
' Class Module 名の有効性を確認する
'
' ClassName と同じ名前の VBA Component が
' すでに存在する場合は生成を中止する。
'
' 特に Worksheet の CodeName との衝突を防止する。
'==================================================
Public Function QDB_ValidateClassName( _
    ByVal ClassName As String) As Boolean

    Dim component As Object

    QDB_ValidateClassName = False

    '------------------------------------------
    ' 空文字チェック
    '------------------------------------------
    If Trim$(ClassName) = "" Then

        MsgBox _
            "Class Module 名が空です。", _
            vbExclamation, _
            "Quick DB"

        Exit Function

    End If

    '------------------------------------------
    ' VBA Component の存在確認
    '------------------------------------------
    On Error Resume Next

    Set component = _
        ThisWorkbook.VBProject.VBComponents( _
            ClassName)

    On Error GoTo 0

    '------------------------------------------
    ' 同名 Component が存在しない
    '------------------------------------------
    If component Is Nothing Then

        QDB_ValidateClassName = True

        Exit Function

    End If

    '------------------------------------------
    ' 同名 Component が存在する
    '------------------------------------------
    MsgBox _
        "Class Module [" & _
        ClassName & _
        "] は使用できません。" & _
        vbCrLf & vbCrLf & _
        "同じ名前の VBA オブジェクトが" & _
        "すでに存在します。" & _
        vbCrLf & vbCrLf & _
        "Worksheet の CodeName などと" & _
        "名前が重複していないか確認してください。", _
        vbExclamation, _
        "Quick DB"

    Debug.Print _
        "ClassName Conflict : " & _
        ClassName

End Function

Public Sub QDB_Generate()

    Dim ws As Worksheet
    Dim ClassName As String
    Dim fieldCount As Long


    Debug.Print "========================================"
    Debug.Print "Quick DB Generate Start"
    Debug.Print "========================================"


    ' --------------------------------------------------------
    ' 対象シート取得
    ' --------------------------------------------------------

    On Error Resume Next

    Set ws = _
        ThisWorkbook.Worksheets(QDB_SHEET_NAME)

    On Error GoTo 0


    If ws Is Nothing Then

        MsgBox _
            "対象シート [" & _
            QDB_SHEET_NAME & _
            "] が見つかりません。", _
            vbExclamation

        Debug.Print _
            "対象シートが見つかりません。"

        Exit Sub

    End If


    ' --------------------------------------------------------
    ' Class名
    '
    ' 現在は旧Quick DBと同じく
    ' シート名をClass名として使用する。
    ' --------------------------------------------------------

    ClassName = ws.Name


    ' --------------------------------------------------------
    ' Schema取得
    ' --------------------------------------------------------

    fieldCount = _
        QDB_GetFieldCount(ws)


    Debug.Print _
        "Sheet       : " & ws.Name

    Debug.Print _
        "Class Name  : " & ClassName

    Debug.Print _
        "Field Count : " & fieldCount


    If fieldCount = 0 Then

        MsgBox _
            "フィールドが見つかりません。", _
            vbExclamation

        Debug.Print _
            "フィールドがありません。"

        Exit Sub

    End If


    ' --------------------------------------------------------
    ' Class生成
    ' --------------------------------------------------------

    If QDB_GENERATE_CLASS Then

        Debug.Print _
            "Generate Class..."

        QDB_GenerateClass _
            ws, _
            ClassName, _
            fieldCount

    End If


    ' --------------------------------------------------------
    ' Standard Module生成
    ' --------------------------------------------------------

    If QDB_GENERATE_MODULE Then

        Debug.Print _
            "Generate Module..."

        QDB_GenerateModule _
            ws, _
            ClassName, _
            fieldCount

    End If


    ' --------------------------------------------------------
    ' 完了
    ' --------------------------------------------------------

    Debug.Print _
        "========================================"

    Debug.Print _
        "Quick DB Generate Complete"

    Debug.Print _
        "========================================"


    MsgBox _
        "Quick DB の生成が完了しました。" & _
        vbCrLf & vbCrLf & _
        "Class : " & ClassName & _
        vbCrLf & _
        "Fields: " & fieldCount, _
        vbInformation, _
        "Quick DB"


End Sub

' ============================================================
' Quick DB Engine
' QDB_PrepareGenerate
'
' Generator実行前の準備処理。
'
' 既存の生成対象Module / Class Moduleが存在する場合、
' それらを削除してから新しいDBを生成する。
'
' ============================================================

Public Sub QDB_PrepareGenerate( _
    ByVal ClassName As String)

    Dim ModuleName As String

    ModuleName = QDB_MODULE_PREFIX & ClassName

    Debug.Print "----------------------------------------"
    Debug.Print "Prepare Generate"
    Debug.Print "Class Module : " & ClassName
    Debug.Print "DB Module    : " & ModuleName
    Debug.Print "----------------------------------------"

    ' --------------------------------------------------------
    ' 既存DB Moduleを削除
    ' --------------------------------------------------------

    QDB_DeleteVBComponent ModuleName

    ' --------------------------------------------------------
    ' 既存Class Moduleを削除
    ' --------------------------------------------------------

    QDB_DeleteVBComponent ClassName

    Debug.Print "Prepare Generate Complete"

End Sub


Private Sub QDB_DeleteVBComponent( _
    ByVal ComponentName As String)

    Dim component As Object

    Set component = QDB_GetVBComponent(ComponentName)

    If component Is Nothing Then

        Debug.Print _
            "  Not Found : " & ComponentName

        Exit Sub

    End If

    On Error GoTo DeleteError

    ThisWorkbook.VBProject.VBComponents.Remove component

    Debug.Print _
        "  Deleted : " & ComponentName

    Exit Sub

DeleteError:

    Debug.Print _
        "  Delete Error : " & _
        ComponentName & _
        " / " & _
        Err.Number & _
        " / " & _
        Err.Description

    MsgBox _
        "Module [" & ComponentName & _
        "] を削除できませんでした。" & vbCrLf & vbCrLf & _
        Err.Number & " / " & Err.Description, _
        vbExclamation

End Sub


Private Function QDB_GetVBComponent( _
    ByVal ComponentName As String) As Object

    On Error Resume Next

    Set QDB_GetVBComponent = _
        ThisWorkbook.VBProject.VBComponents(ComponentName)

    On Error GoTo 0

End Function

