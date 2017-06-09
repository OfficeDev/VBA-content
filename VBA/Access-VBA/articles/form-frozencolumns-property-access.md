---
title: Form.FrozenColumns Property (Access)
keywords: vbaac10.chm13420
f1_keywords:
- vbaac10.chm13420
ms.prod: access
api_name:
- Access.Form.FrozenColumns
ms.assetid: 5b595c5e-6a2e-e3d8-1ae8-a2f224eb5516
ms.date: 06/08/2017
---


# Form.FrozenColumns Property (Access)

You can use the  **FrozenColumns** property to determine how many columns in a datasheet are frozen. Read/write **Integer**.


## Syntax

 _expression_. **FrozenColumns**

 _expression_ A variable that represents a **Form** object.


## Remarks

Frozen columns are displayed on the left side of the datasheet and don't move when you scroll horizontally through the datasheet.


 **Note**  The  **FrozenColumns** property applies only to tables, forms, and queries in Datasheet view.

In [Visual Basic](set-properties-by-using-visual-basic.md), this property setting is an  **Integer** value indicating the number of columns in the datasheet that have been frozen by using the **Freeze Columns** command. The record selector column is always frozen, so the default value is 1. Consequently, if you freeze one column, the **FrozenColumns** property is set to 2; if you freeze two columns, it's set to 3; and so on.


## Example

The following example uses the  **FrozenColumns** property to determine how many columns are frozen in a table in Datasheet view. If more than three columns are frozen, the table size is maximized so you can see as many unfrozen columns as possible.


```vb
Sub CheckFrozen(strTableName As String) 
 Dim dbs As Object 
 Dim tdf As Object 
 Dim prp As Variant 
 Const DB_Integer As Integer = 3 
 Const conPropertyNotFound = 3270 ' Property not found error. 
 Set dbs = CurrentDb ' Get current database. 
 Set tdf = dbs.TableDefs(strTableName) ' Get object for table. 
 DoCmd.OpenTable strTableName, acNormal ' Open table. 
 tdf.Properties.Refresh 
 On Error GoTo Frozen_Err 
 If tdf.Properties("FrozenColumns") > 3 Then ' Check property. 
 DoCmd.Maximize 
 End If 
Frozen_Bye: 
 Exit Sub 
Frozen_Err: 
 If Err = conPropertyNotFound Then ' Property not in collection. 
 Set prp = tdf.CreateProperty("FrozenColumns", DB_Integer, 1) 
 tdf.Properties.Append prp 
 Resume Frozen_Bye 
 End If 
End Sub
```


## See also


#### Concepts


[Form Object](form-object-access.md)

