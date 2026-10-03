# VBA Quick DB Engine

Excel worksheet data to VBA Class Modules and data modules.

**VBA Quick DB Engine** is a small VBA engine that generates VBA Class Modules and data modules from structured data defined in an Excel worksheet.

It is not intended to be a general-purpose database management system.

The main idea is simple:

```text
Excel Worksheet
      ↓
Table definition + data
      ↓
Quick DB Engine
      ↓
VBA Class Module
      +
VBA Data Module
      ↓
Use from your own VBA code
```

---

## Overview

Quick DB started as a small experiment for handling master data in Excel/VBA.

The original idea was to write structured data in an Excel worksheet and generate the corresponding VBA code automatically.

For example, a worksheet can define:

```text
        E       F       G       H
    ┌───────┬───────┬───────┬───────┐
 5  │  ID   │   X   │   Y   │  Val  │
    ├───────┼───────┼───────┼───────┤
 6  │String │ Long  │ Long  │String │
    ├───────┼───────┼───────┼───────┤
 7  │ 1000  │  10   │  15   │   A   │
 8  │ 1001  │  20   │  25   │   B   │
 9  │ 1002  │  30   │  35   │   C   │
10  │ 1003  │  50   │  55   │   D   │
    └───────┴───────┴───────┴───────┘
```

Quick DB reads this definition and generates VBA code such as:

```vb
Public ID As String
Public X As Long
Public Y As Long
Public Val As String
```

and a data initialization routine such as:

```vb
data.Add CreateData_ClassData("1000", 10, 15, "A")
data.Add CreateData_ClassData("1001", 20, 25, "B")
data.Add CreateData_ClassData("1002", 30, 35, "C")
data.Add CreateData_ClassData("1003", 50, 55, "D")
```

The Excel worksheet therefore becomes the place where the structure and data are designed.

---

## What Quick DB Actually Does

The name "DB" may suggest a full database system.

Quick DB is much smaller than that.

The core function is:

```text
Worksheet
   ↓
Schema
   ↓
Class Module
   +
Data Module
```

It does not try to replace a relational database system.

It does not provide a complete SQL-like environment, transaction system, indexing system, or full CRUD framework.

The generated data is intended to become a part of another VBA application.

For example:

```vb
Call init_ClassData

' Use result_ClassData here
```

After generation, the developer can write their own search, filtering, sorting, game logic, business logic, or other processing.

---

## Why Excel?

This project follows a simple idea:

> Write the design in an Excel worksheet, then let another mechanism process it.

Excel is useful here because it provides a familiar grid for defining structured data.

The worksheet is not necessarily the final runtime data store.

It can instead act as a design surface.

```text
Human
  ↓
Excel Worksheet
  ↓
Generator
  ↓
VBA Code
  ↓
Application
```

This is one example of the broader idea behind **The Excel Plate**.

---

## The Excel Plate

The Excel Plate is a concept used across several of my Excel/VBA projects.

The basic idea is:

> **Design it in a worksheet, then make another system do the work.**

Excel does not have to perform every operation itself.

It can be used as a "plate" on which structured definitions are placed.

Quick DB is one example:

```text
Excel
  ↓
Data / structure definition
  ↓
Quick DB
  ↓
VBA source code
```

Other projects can use the same basic approach with completely different targets.

The important part is not the particular tool.

The important part is using the Excel worksheet as a practical design surface.

---

## Example

The sample workbook contains a worksheet named:

```text
ClassData
```

The table starts at cell `E5`.

### Field definition

Row 5 contains field names.

```text
ID    X    Y    Val
```

### Type definition

Row 6 contains the VBA data types.

```text
String    Long    Long    String
```

### Data

Data starts from row 7.

```text
1000    10    15    A
1001    20    25    B
1002    30    35    C
1003    50    55    D
```

The generated class is based on the worksheet name.

Therefore:

```text
ClassData
    ↓
ClassData.cls
```

The generated standard module is named:

```text
QDB_ClassData
```

---

## Generated Class Module

For the sample table, Quick DB generates a Class Module similar to:

```vb
Option Explicit

Public ID As String
Public X As Long
Public Y As Long
Public Val As String

Public Sub init_ClassData( _
    ID As String, _
    X As Long, _
    Y As Long, _
    Val As String)

    Me.ID = ID
    Me.X = X
    Me.Y = Y
    Me.Val = Val

End Sub
```

The generated class represents one record.

---

## Generated Data Module

Quick DB also generates a standard module.

The module contains the generated data and creates instances of the generated Class Module.

Conceptually:

```text
ClassData
    ↓
one ClassData object = one record

QDB_ClassData
    ↓
Collection of ClassData objects
```

The generated module provides:

```vb
result_ClassData
```

and:

```vb
ClassData_Field
```

for accessing the generated data and record count.

---

## Searching the Data

The repository currently contains Finder and View modules as examples of how generated data can be processed.

For example:

```vb
Call QDB_DataView( _
    result_ClassData, _
    field_name, _
    field_value, _
    operator_type, _
    FieldNames)
```

These functions are not intended to define a complete database query language.

They are simply examples of what can be done with the generated data.

A developer using Quick DB can replace or extend these routines with their own functions.

For example:

```text
Quick DB
    ↓
Generate data
    ↓
Your application
    ├─ Find
    ├─ Filter
    ├─ Sort
    ├─ Group
    └─ Game / business logic
```

The generated data is the important part.

The way that data is used is left open.

---

## Project Structure

The repository contains the following VBA source modules.

```text
src
├─ QDB_01_Config.bas
├─ QDB_10_Schema.bas
├─ QDB_20_Generator.bas
├─ QDB_30_ClassGenerator.bas
├─ QDB_40_ModuleGenerator.bas
├─ QDB_50_Finder.bas
├─ QDB_60_View.bas
└─ QDB_90_Test.bas
```

### QDB_01_Config

Contains configuration constants used by Quick DB.

Examples include the worksheet name and table position.

```text
Field row       : 5
Type row        : 6
Data start row  : 7
Table start col : E
```

### QDB_10_Schema

Reads the worksheet schema.

It determines:

* field count
* field names
* field types
* data range

### QDB_20_Generator

Controls the generation process.

It connects the schema information with the Class and Module generators.

### QDB_30_ClassGenerator

Generates the VBA Class Module.

The generated class is based on the worksheet name.

### QDB_40_ModuleGenerator

Generates the standard data module.

It creates the data initialization code from the worksheet rows.

### QDB_50_Finder

Provides simple search functions for generated data.

This is an example implementation rather than a complete database query engine.

### QDB_60_View

Provides a simple way to display search results.

It is primarily intended for testing and demonstration.

### QDB_90_Test

Contains example usage and the Struct Mode cleanup routine.

---

## Struct Mode and Run Mode

Quick DB can be thought of as having two stages.

### Struct Mode

The Builder modules are present.

The developer prepares the worksheet and generates the VBA code.

```text
Worksheet
    ↓
Generate
    ↓
Class Module
Data Module
```

### Run Mode

After generation, the Builder modules are no longer necessary for normal use.

The generated Class Module and data module remain.

The application can then use the generated data directly.

Conceptually:

```text
STRUCT MODE

Builder
   ↓
Generate
   ↓
Generated VBA


RUN MODE

Generated VBA
   ↓
Your application
```

The repository includes a routine for removing the Builder modules after generation.

---

## Important: VBA Project Access

Quick DB generates VBA components by accessing the workbook's VBA project.

It uses the Excel VBA extensibility model through:

```vb
ThisWorkbook.VBProject.VBComponents
```

Therefore, Excel must allow programmatic access to the VBA project.

Depending on the Excel security settings, the following Trust Center option may be required:

**Trust access to the VBA project object model**

Without this permission, Quick DB cannot create, remove, or import VBA components programmatically.

This is an important part of the design and is not required by ordinary worksheet-processing macros.

---

## Component Name Restrictions

The generated Class Module uses the worksheet name.

For example:

```text
Worksheet name
    ↓
Class name
```

Therefore, the worksheet name must not conflict with an existing VBA component.

For example, the default worksheet CodeName:

```text
Sheet1
```

can already exist as a VBA component.

Using `Sheet1` as the generated Class name would therefore cause a naming collision.

Quick DB validates the component name before generation.

A worksheet such as:

```text
ClassData
```

is suitable for the sample.

---

## Regeneration

Quick DB regenerates the generated Class and data modules.

The intended workflow is:

```text
Modify worksheet
      ↓
Generate again
      ↓
Old generated modules are replaced
      ↓
New VBA code is created
```

This means that generated modules should be treated as generated output.

Do not put custom application code into the generated modules if that code needs to survive regeneration.

Instead, keep application logic in separate modules.

---

## Basic Workflow

1. Open the sample workbook.
2. Prepare the worksheet table.
3. Check the field names and type definitions.
4. Run the Quick DB generation routine.
5. The Class Module and data module are generated.
6. Initialize the generated data.
7. Use the generated Collection from your own VBA code.
8. Optionally remove the Builder modules when Struct Mode is finished.

The generated code is ordinary VBA.

There is no special runtime library required after generation.

---

## Development Environment

The original development environment for this project is:

```text
Microsoft Excel 2007
VBA 6.5
Windows 7
32-bit
```

The code was designed with older VBA compatibility in mind.

It does not depend on modern VBA syntax such as `PtrSafe`.

---

## Relationship with vba-develop-boosters

Quick DB and `vba-develop-boosters` solve different problems.

`vba-develop-boosters` is concerned with moving VBA components into a workbook.

Quick DB is concerned with generating and managing VBA components from worksheet definitions.

Conceptually:

```text
Source / Repository / Workbook
            ↓
   vba-develop-boosters
          [搬送]
            ↓
        Quick DB
          [生成]
            ↓
     VBA Components
            ↓
      Your Workbook
```

These projects can therefore complement each other.

---

## Background

Quick DB was originally created as an experiment for handling RPG-related master data in Excel/VBA.

The original project was much more focused on the idea of quickly creating structured data from an Excel worksheet.

Over time, the idea became more general:

```text
Excel worksheet
       ↓
   schema definition
       ↓
   generated class
       ↓
   generated data
```

The current repository is a cleaned-up version of that idea.

The purpose of publishing it is not to present a full database product.

It is to make the technique itself available.

---

## What This Project Is Not

Quick DB is not intended to be:

* a relational database management system
* a replacement for SQL
* a complete ORM
* a full CRUD framework
* a general-purpose data access layer
* a high-performance database engine

It is a small VBA code-generation engine.

The interesting part is the connection between:

```text
Excel worksheet design
        ↓
VBA source generation
        ↓
VBA application
```

---

## Future Possibilities

There are many possible directions for this type of system.

For example:

* additional VBA data types
* more sophisticated validation
* additional generated methods
* custom search functions
* sorting and filtering helpers
* dependency management
* component assembly
* integration with other VBA development tools

However, these are possibilities rather than promises.

The current project intentionally remains small.

---

## Final Thought

Excel is often treated only as the place where the final data is entered.

This project explores another possibility.

```text
Excel
  ↓
Design
  ↓
Generate
  ↓
Code
  ↓
Run
```

The worksheet can be more than a spreadsheet.

It can be a design surface.

Quick DB is one small experiment built around that idea.

---

## License

No license is currently specified for this repository.
