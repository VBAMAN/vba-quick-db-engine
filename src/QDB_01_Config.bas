Attribute VB_Name = "QDB_01_Config"
Option Explicit

' ============================================================
' Quick DB Engine
' QDB_01_Config
'
' システム全体で使用する設定値
' ============================================================

' ------------------------------------------------------------
' 対象シート
' ------------------------------------------------------------

Public Const QDB_SHEET_NAME As String = "ClassData"


' ------------------------------------------------------------
' シート構造
'
' 1行目 : フィールド名
' 2行目 : VBAデータ型
' 3行目以降 : データ
' ------------------------------------------------------------

Public Const QDB_FIELD_ROW As Long = 5
Public Const QDB_TYPE_ROW As Long = 6
Public Const QDB_DATA_START_ROW As Long = 7

Public Const QDB_TABLE_COL As Long = 5

' ------------------------------------------------------------
' フィールド数の上限
' ------------------------------------------------------------

Public Const QDB_MAX_FIELDS As Long = 256


' ------------------------------------------------------------
' 生成されるVBAコンポーネント名
' ------------------------------------------------------------

Public Const QDB_CLASS_MODULE As String = "CLASS"
Public Const QDB_STANDARD_MODULE As String = "STANDARD"


' ------------------------------------------------------------
' 生成モジュール名の接頭辞
'
' 例：
'   ClassData
'       ↓
'   mod_ClassData
' ------------------------------------------------------------

Public Const QDB_MODULE_PREFIX As String = "QDB_"


' ------------------------------------------------------------
' IDフィールド
'
' 現段階では旧Quick DBとの互換性を維持するため使用。
' 将来的にはID固定を解除する予定。
' ------------------------------------------------------------

Public Const QDB_ID_FIELD As String = "ID"


' ------------------------------------------------------------
' Generatorの動作設定
' ------------------------------------------------------------

Public Const QDB_GENERATE_CLASS As Boolean = True
Public Const QDB_GENERATE_MODULE As Boolean = True

