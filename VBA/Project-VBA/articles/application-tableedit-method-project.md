---
title: Application.TableEdit Method (Project)
keywords: vbapj.chm403
f1_keywords:
- vbapj.chm403
ms.prod: project-server
api_name:
- Project.Application.TableEdit
ms.assetid: 370ab75d-9b99-b4b3-db5f-96697320bc68
ms.date: 06/08/2017
---


# Application.TableEdit Method (Project)

Creates, edits, or copies a table.


## Syntax

 _expression_. **TableEdit**( ** _Name_**, ** _TaskTable_**, ** _Create_**, ** _OverwriteExisting_**, ** _NewName_**, ** _FieldName_**, ** _NewFieldName_**, ** _Title_**, ** _Width_**, ** _Align_**, ** _ShowInMenu_**, ** _LockFirstColumn_**, ** _DateFormat_**, ** _RowHeight_**, ** _ColumnPosition_**, ** _AlignTitle_**, ** _HeaderAutoRowHeightAdjustment_**, ** _HeaderTextWrap_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**| The name of a table to edit, create, or copy.|
| _TaskTable_|Required|**Boolean**|**True** if the active table contains information about tasks or resources; otherwise, **False**.|
| _Create_|Optional|**Boolean**|**True** if Project creates a table, otherwise, **False**. If NewName is not defined, the new table is given the name specified with Name. Otherwise, the new table is a copy of the table specified with Name and is given the name specified with NewName. The default value is **False**.|
| _OverwriteExisting_|Optional|**Boolean**|**True** if an existing table is overwritten with the new table. The default value is **False**.|
| _NewName_|Optional|**String**|The new name for the existing table (Create is  **False** ) or new table (Create is **True** ). If NewName is not defined and Create is **False**, the table specified with Name retains its current name. The default value is an empty string ("").|
| _FieldName_|Optional|**String**|The name of a field to change.|
| _NewFieldName_|Optional|**String**|The name of a new field. The field specified with NewFieldName replaces the field specified with FieldName.|
| _Title_|Optional|**String**|The title for the field specified with FieldName.|
| _Width_|Optional|**Integer**|A number that specifies the width of the field specified with FieldName. The default value is 10 for new fields.|
| _Align_|Optional|**Integer**|Specifies how to align the text in the field specified with FieldName. Can be one of the following  **[PjAlignment](pjalignment-enumeration-project.md)** constants: **pjLeft**, **pjCenter**, or **pjRight**. The default value is **pjRight**.|
| _ShowInMenu_|Optional|**Boolean**|**True** if the table name appears in the **Tables** drop-down menu; otherwise, **False**. (The **Tables** drop-down menu is on the **View** tab of the Ribbon.) The default value is **False.**|
| _LockFirstColumn_|Optional|**Boolean**|**True** if Project locks or prevents changes to the first column of the table; otherwise, **False**. The default value is **False**.|
| _DateFormat_|Optional|**Integer**|A constant that specifies the format for the date fields in the table. Can be one of the  **[PjDateFormat](pjdateformat-enumeration-project.md)** constants. The default value is **pjDateDefault**.|
| _RowHeight_|Optional|**Integer**|The height of the rows in the table. The default value is 1.|
| _ColumnPosition_|Optional|**Long**|The number of the column to edit. (Columns are numbered from left to right, starting with 0.) If a value for NewFieldName is specified, a new column is inserted in the table. If ColumnPosition is set to 0, the new field is inserted in the first column (LockFirstColumn is  **False** ) or the second column (LockFirstColumn is **True** ) of the table. Set ColumnPosition to -1 to specify the last column of the table. The default value is -1.|
| _AlignTitle_|Optional|**Long**|A constant that specifies the alignment of the column title. Can be one of the following  **PjAlignment** constants: **pjLeft**, **pjCenter**, or **pjRight**. The default value is **pjCenter**.|
| _HeaderAutoRowHeightAdjustment_|Optional|**Boolean**|**True** if Project automatically adjusts the row height of the table; otherwise, **False**. The default value is **True**.|
| _HeaderTextWrap_|Optional|**Boolean**|**True** if Project wraps text in the header of the table; otherwise, **False**. The default value is **True**.|

### Return Value

 **Boolean**


## Remarks

Project sets the order of years, months, and days in a date format equal to the corresponding value in the  **Regional and Language Options** dialog box of the Windows Control Panel.

To make a copy of the active table, see the  **[TableCopy](application-tablecopy-method-project.md)** method. To include options to wrap text within the table and use the **Add New Column** feature, see the **[TableEditEx](application-tableeditex-method-project.md)** method.


## Example

The following example creates a new table based on the Task Usage table and adds the table to the  **Table** drop-down menu. The macro adds the Priority field as the second column with a title and width of 12, changes the default date format, and then makes the new table the active view.


```vb
Sub CreateNewTaskUsageTable() 
 TableEdit Name:="Usage", TaskTable:=True, Create:=True, _ 
 NewName:="Priority Tasks" 
 
 TableEdit Name:="Priority Tasks", TaskTable:=True, _ 
 NewFieldName:="Priority", Title:="Priority", Width:=12, _ 
 ShowInMenu:=True, DateFormat:=pjDate_mm_dd_yy, _ 
 ColumnPosition:=1 
 
 TableApply "Priority Tasks" 
End Sub
```


