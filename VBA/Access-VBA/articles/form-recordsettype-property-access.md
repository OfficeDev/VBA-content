---
title: Form.RecordsetType Property (Access)
keywords: vbaac10.chm13361
f1_keywords:
- vbaac10.chm13361
ms.prod: access
ms.assetid: 29690204-1014-961d-a969-25c44ca5fc6e
ms.date: 06/08/2017
---


# Form.RecordsetType Property (Access)

You can use the  **RecordsetType** property to specify what kind of recordset is made available to a form. Read/write **Byte**.


## Syntax

 _expression_. **RecordsetType**

 _expression_ A variable that represents an **Form** object.


## Remarks

The  **RecordsetType** property uses the following settings in a Microsoft Access database.



|**Setting**|**Type of Recordset**|**Description**|
|:-----|:-----|:-----|
|0|Dynaset|(Default) You can edit bound controls based on a single table or tables with a one-to-one relationship. For controls bound to fields based on tables with a one-to-many relationship, you can't edit data from the join field on the "one" side of the relationship unless cascade update is enabled between the tables|
|1|Dynaset (Inconsistent Updates)|All tables and controls bound to their fields can be edited.|
|2|Snapshot|No tables or the controls bound to their fields can be edited.|

 **Note**  If you don't want data in bound controls to be edited when a form is in Form view or Datasheet view, you can set the  **RecordsetType** property to 2.


 **Note**  Changing the  **RecordsetType** property of an open form or report causes an automatic recreation of the recordset.

You can create forms based on multiple underlying tables with fields bound to controls on the forms. Depending on the  **RecordsetType** property setting, you can limit which of these bound controls can be edited.

In addition to the editing control provided by  **RecordsetType**, each control on a form has a **Locked** property that you can set to specify whether the control and its underlying data can be edited. If the **Locked** property is set to Yes, you can't edit the data.


## Example

In the following example, only if the user ID is ADMIN can records be updated. This code sample sets the  **RecordsetType** property to Snapshot if the public variable gstrUserID value is not ADMIN.


```vb
Sub Form_Open(Cancel As Integer) 
 Const conSnapshot = 2 
 If gstrUserID <> "ADMIN" Then 
 Forms!Employees.RecordsetType = conSnapshot 
 End If 
End Sub
```


## Property value

 **UINT8**


