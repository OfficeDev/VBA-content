---
title: Application.TableEditEx Method (Project)
keywords: vbapj.chm2172
f1_keywords:
- vbapj.chm2172
ms.prod: project-server
api_name:
- Project.Application.TableEditEx
ms.assetid: 953cdbf6-24ac-5e39-9c23-ec05ec9e4809
ms.date: 06/08/2017
---


# Application.TableEditEx Method (Project)

Creates, edits, or copies a table that can wrap text and include the  **Add New Column** feature.


## Syntax

 _expression_. **TableEditEx**( ** _Name_**, ** _TaskTable_**, ** _Create_**, ** _OverwriteExisting_**, ** _NewName_**, ** _FieldName_**, ** _NewFieldName_**, ** _Title_**, ** _Width_**, ** _Align_**, ** _ShowInMenu_**, ** _LockFirstColumn_**, ** _DateFormat_**, ** _RowHeight_**, ** _ColumnPosition_**, ** _AlignTitle_**, ** _HeaderAutoRowHeightAdjustment_**, ** _HeaderTextWrap_**, ** _WrapText_**, ** _ShowAddNewColumn_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**| The name of a table to edit, create, or copy.|
| _TaskTable_|Required|**Boolean**|**True** if the active table contains information about tasks or resources; otherwise, **False**.|
| _Create_|Optional|**Boolean**|**True** if Project creates a table; otherwise, **False**. If _NewName_ is not defined, the new table is given the name specified by _Name_. Otherwise, the new table is a copy of the table specified by  _Name_ and is given the name specified by _NewName_. The default value is  **False**.|
| _OverwriteExisting_|Optional|**Boolean**|**True** if an existing table is overwritten with the new table; otherwise, **False**. The default value is **False**.|
| _NewName_|Optional|**String**|The new name for the existing table ( _Create_ is **False** ) or new table ( _Create_ is **True** ). If _NewName_ is not defined and _Create_ is **False**, the table specified by _Name_ retains its current name. The default value is an empty string ("").|
| _FieldName_|Optional|**String**|The name of a field to change.|
| _NewFieldName_|Optional|**String**|The name of a new field. The field specified by  _NewFieldName_ replaces the field specified by _FieldName_.|
| _Title_|Optional|**String**|The title for the field specified by  _FieldName_.|
| _Width_|Optional|**Integer**|A number that specifies the width of the field specified by  _FieldName_. The default value is 10 for new fields.|
| _Align_|Optional|**Integer**|Specifies how to align the text in the field specified by  _FieldName_. Can be one of the following  **[PjAlignment](pjalignment-enumeration-project.md)** constants: **pjLeft**, **pjCenter**, or **pjRight**. The default value is **pjRight**.|
| _ShowInMenu_|Optional|**Boolean**|**True** if the table name appears in the **Tables** drop-down menu; otherwise, **False**. (The **Tables** drop-down menu is on the **VIEW** ribbon.) The default value is **False.**|
| _LockFirstColumn_|Optional|**Boolean**|**True** if Project locks or prevents changes to the first column of the table; otherwise, **False**. The default value is **False**.|
| _DateFormat_|Optional|**Integer**|A constant that specifies the format for the date fields in the table. Can be one of the  **[PjDateFormat](pjdateformat-enumeration-project.md)** constants. The default value is **pjDateDefault**.|
| _RowHeight_|Optional|**Integer**|The height of the rows in the table. The default value is 1.|
| _ColumnPosition_|Optional|**Long**|The number of the column to edit. (Columns are numbered from left to right, starting with 0.) If  _NewFieldName_ is specified, a new column is inserted in the table. If _ColumnPosition_ is set to 0, the new field is inserted in the first column ( _LockFirstColumn_ is **False** ) or the second column ( _LockFirstColumn_ is **True** ) of the table. Set _ColumnPosition_ to -1 to specify the last column of the table. The default value is -1.|
| _AlignTitle_|Optional|**Long**|A constant that specifies the alignment of the column title. Can be one of the following  **PjAlignment** constants: **pjLeft**, **pjCenter**, or **pjRight**. The default value is **pjCenter**.|
| _HeaderAutoRowHeightAdjustment_|Optional|**Boolean**|**True** if Project automatically adjusts the row height of the table; otherwise, **False**. The default value is **True**.|
| _HeaderTextWrap_|Optional|**Boolean**|**True** if Project wraps text in the header of the table; otherwise, **False**. The default value is **True**.|
| _WrapText_|Optional|**Boolean**|**True** if the table wraps text in the rows; otherwise, **False**.|
| _ShowAddNewColumn_|Optional|**Boolean**|True if the table shows the  **Add New Column** feature in the far-right column; otherwise, **False**.|

### Return Value

 **Boolean**


## Remarks

Project sets the order of years, months, and days in a date format equal to the corresponding value in the  **Regional and Language Options** dialog box of the Windows Control Panel.

To make a copy of the active table, see the  **[TableCopy](application-tablecopy-method-project.md)** method.


## Example

The following example creates a table based on the Task Usage table, includes the  **Add New Column** feature, and adds the table to the **Table** drop-down menu. The macro adds the Priority field as the second column with a title and width of 12, changes the default date format, and then makes the new table the active view.


```vb
Sub CreateNewTaskUsageTable() 
    TableEditEx Name:="Usage", TaskTable:=True, Create:=True, _ 
        NewName:="Priority Tasks", ShowAddNewColumn:=True 
 
    TableEditEx Name:="Priority Tasks", TaskTable:=True, _ 
        NewFieldName:="Priority", Title:="Priority", Width:=12, _ 
        ShowInMenu:=True, DateFormat:=pjDate_mm_dd_yy, _ 
        ColumnPosition:=1 
 
    TableApply "Priority Tasks" 
End Sub
```


