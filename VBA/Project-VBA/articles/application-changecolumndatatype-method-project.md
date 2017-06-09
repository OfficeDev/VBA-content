---
title: Application.ChangeColumnDataType Method (Project)
keywords: vbapj.chm711
f1_keywords:
- vbapj.chm711
ms.prod: project-server
api_name:
- Project.Application.ChangeColumnDataType
ms.assetid: 25cbcb73-4cbd-3ea7-ff16-90a4d3028af9
ms.date: 06/08/2017
---


# Application.ChangeColumnDataType Method (Project)

Changes the data type of a local custom field column in a table.


## Syntax

 _expression_. **ChangeColumnDataType**( ** _Type_**, ** _Column_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required|**PjFieldTypes**|Specifies the type of the custom field data. The value can be one of the  **[PjFieldTypes](pjfieldtypes-enumeration-project.md)** constants. The default value is 0 ( **pjCostField** ).|
| _Column_|Optional|**Variant**|Specifies the absolute column location. A value of 0 changes the data type of a column in the left-most position, if that column is a local custom field. If the first column is locked, the left-most position is the first column after the locked column. The default value is the selected column.|

### Return Value

 **Boolean**


## Remarks

 **ChangeColumnDataType** requires a custom field column to be selected. To manually change the data type of a custom field column, add a custom field column to a table in a view, right-click the column header, and then click **Data Type**.


## Example

To use the following example, create a project with several tasks, and then open the Gantt Chart view. The  **CreateTestTable** macro creates a task table that has four columns. The first column with the ID field is locked. The second column has the title **Task Name**, the third column contains the  **Text1** task custom field, and the fourth column contains the **Number1** custom field. The macro assigns the table to the current view, and then adds text and number values to the task custom fields.




1. Run the  **CreateTestTable** macro. The **Text1** custom field value of the first task is **42 X**.
    
2. Run the  **SwitchNumberAndText** macro. The macro switches the headings and types of the two custom fields.
    
     **Note**  Because the value of the  **Text1** custom field in the first task is **42 X**, when  **ChangeColumnDataType** tries to convert that column to the **Number1** custom field, Project shows an error dialog box with the message, **Converting this data will cause errors. The contents of 1 records will be deleted. Do you want to continue anyway?**
3. To continue with the conversion, click  **Yes** in the error dialog box. When the **Text1** custom field changes to the **Number1** custom field, the value **42 X** changes to **0**.
    
4. To change back to a standard table in the Gantt Chart view, right-click the  **Select All** cell (the unnamed top-left cell in the table), and then select a different table in the drop-down list.
    





```vb
Sub CreateTestTable() 
    Dim t As Task 
    Dim n As Integer 
 
    TableEditEx Name:="Task Test Table", TaskTable:=True, Create:=True, FieldName:="ID", _ 
        Width:=5, ShowInMenu:=True, HeaderAutoRowHeightAdjustment:=True, _ 
    ShowAddNewColumn:=False 
 
    TableEditEx Name:="Task Test Table", TaskTable:=True, NewFieldName:="Name", Title:="Task Name" 
    TableEditEx Name:="Task Test Table", TaskTable:=True, NewFieldName:="Text1" 
    TableEditEx Name:="Task Test Table", TaskTable:=True, NewFieldName:="Number1" 
    TableEditEx Name:="Task Test Table", TaskTable:=True, LockFirstColumn:=True 
 
    TableApply Name:="Task Test Table" 
 
    n = 42 

    For Each t In ActiveProject.Tasks 
        If n = 42 Then 
            t.Text1 = CStr(n) &; " X" 
        Else 
            t.Text1 = CStr(n) 
        End If 
 
        t.Number1 = n 
        n = n + 2 
    Next t 
End Sub 
 
Sub SwitchNumberAndText() 
    SelectTaskColumn Column:="Number1" 
    ChangeColumnDataType Type:=pjTextField 
 
    SelectTaskColumn Column:="Text1" 
    ChangeColumnDataType Type:=pjNumberField 
End Sub
```


