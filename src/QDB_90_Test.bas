Attribute VB_Name = "QDB_90_Test"
Option Explicit

'============================================================
' Quick DB Engine
' Test Module
'
' Quick DBで生成したデータを使用して、
' Finder / View の動作を確認するためのテストコード。
'
' このモジュールは、Quick DB本体が生成した
' Class Module / Standard Moduleとは異なり、
' テストおよび使用例を記述するためのモジュール。
'
' Struct ModeでQuick DBを構築した後、
' Run Modeでは不要であれば手動で削除してよい。
'============================================================

'============================================================
' 数値フィールド検索テスト
'
' ClassData の X フィールドを対象に、
'
'     X >= 30
'
' の条件で検索する。
'
' 検索結果として、
' ID / X / Y の3フィールドを表示する。
'
' この処理では、生成された
'
'     result_ClassData
'
' を検索対象として使用する。
'============================================================

Public Sub USE_DB_Number()


Dim field_name As String
Dim field_value As Double
Dim operator_type As String

Dim FieldNames(1 To 3) As String

'--------------------------------------------------------
' 検索条件
'
' X >= 30
'--------------------------------------------------------
field_name = "X"
field_value = 30
operator_type = ">="

'--------------------------------------------------------
' 検索結果として表示するフィールド
'
' これらの名前は、生成されたClassの
' フィールド名と一致させる。
'--------------------------------------------------------
FieldNames(1) = "ID"
FieldNames(2) = "X"
FieldNames(3) = "Y"

'--------------------------------------------------------
' Quick DB Viewを実行
'
' result_ClassData：
'     ClassDataの生成済みデータ
'
' field_name：
'     検索対象フィールド
'
' field_value：
'     検索値
'
' operator_type：
'     比較演算子
'
' FieldNames：
'     表示するフィールド一覧
'--------------------------------------------------------
Call QDB_DataView( _
    result_ClassData, _
    field_name, _
    field_value, _
    operator_type, _
    FieldNames)


End Sub

'============================================================
' 文字列フィールド検索テスト
'
' ClassData の Val フィールドを対象に、
'
'     Val = "B"
'
' の条件で検索する。
'
' 検索結果として、
' ID / X / Y / Val の4フィールドを表示する。
'
' この例では、生成された
'
'     result_ClassData
'
' を検索対象として使用する。
'============================================================

Public Sub USE_DB_STRING()


Dim field_name As String
Dim field_value As String

Dim FieldNames(1 To 4) As String

'--------------------------------------------------------
' 検索条件
'
' Val = "B"
'--------------------------------------------------------
field_name = "Val"  ' 検索対象フィールド
field_value = "B"   ' 検索文字列

'--------------------------------------------------------
' 検索結果として表示するフィールド
'
' これらはシート上で定義したフィールド名と
' 一致させる必要がある。
'--------------------------------------------------------
FieldNames(1) = "ID"
FieldNames(2) = "X"
FieldNames(3) = "Y"
FieldNames(4) = "Val"

'--------------------------------------------------------
' Quick DB Viewを実行
'--------------------------------------------------------
Call QDB_DataView( _
    result_ClassData, _
    field_name, _
    field_value, _
    "=", _
    FieldNames)


End Sub

'============================================================
' Quick DB Engine
' Struct Mode終了処理
'
' Quick DBによるClass / Module生成が完了した後、
' Run Modeでは不要となるBuilder用モジュールを削除する。
'
' 対象：
'
'   QDB_10_Schema
'   QDB_20_Generator
'   QDB_30_ClassGenerator
'   QDB_40_ModuleGenerator
'
' これらはQuick DBの生成処理を担当するモジュールであり、
' 生成されたClass / Data Moduleそのものではない。
'
' 対象モジュールが存在しない場合は、
' エラーにせずそのまま処理を続行する。
'
' ※ この処理は、生成されたDBをRun Modeで使用する前に
'   必要に応じて実行する。
'============================================================

Public Sub QDB_RemoveBuilderModules()


Debug.Print "========================================"
Debug.Print "Remove Builder Modules"
Debug.Print "========================================"

'--------------------------------------------------------
' Builder用モジュールを削除
'--------------------------------------------------------
QDB_DeleteModule "QDB_10_Schema"
QDB_DeleteModule "QDB_20_Generator"
QDB_DeleteModule "QDB_30_ClassGenerator"
QDB_DeleteModule "QDB_40_ModuleGenerator"

Debug.Print "========================================"
Debug.Print "Remove Builder Modules Complete"
Debug.Print "========================================"


End Sub

'============================================================
' 指定したModuleを削除する
'
' Moduleが存在しない場合：
'     何もしない
'
' 削除に失敗した場合：
'     Immediate Windowへエラー情報を出力する
'
' ThisWorkbookのVBProjectを直接操作するため、
' Excelの「VBAプロジェクト オブジェクト モデルへの
' アクセスを信頼する」設定が必要。
'============================================================

Private Sub QDB_DeleteModule( _
ByVal ModuleName As String)


Dim component As Object

'--------------------------------------------------------
' 対象Componentを取得
'--------------------------------------------------------
Set component = _
    QDB_GetVBComponent(ModuleName)

'--------------------------------------------------------
' 存在しない場合
'--------------------------------------------------------
If component Is Nothing Then

    Debug.Print _
        "  Not Found : " & _
        ModuleName

    Exit Sub

End If

'--------------------------------------------------------
' Module削除
'--------------------------------------------------------
On Error GoTo DeleteError

ThisWorkbook.VBProject.VBComponents.Remove _
    component

Debug.Print _
    "  Deleted : " & _
    ModuleName

Exit Sub


DeleteError:


Debug.Print _
    "  Delete Error : " & _
    ModuleName & _
    " / " & _
    Err.Number & _
    " / " & _
    Err.Description


End Sub

'============================================================
' VBA ProjectからComponentを取得する
'
' 指定した名前のComponentが存在しない場合は、
' Nothingを返す。
'
' Object型で受け取ることで、
' Excel VBAのVBComponentsを参照設定に依存せず扱う。
'============================================================

Private Function QDB_GetVBComponent( _
ByVal ModuleName As String) As Object


On Error Resume Next

Set QDB_GetVBComponent = _
    ThisWorkbook.VBProject.VBComponents( _
        ModuleName)

On Error GoTo 0


End Function
