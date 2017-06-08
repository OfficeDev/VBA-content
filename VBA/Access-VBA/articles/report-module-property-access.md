---
title: Report.Module Property (Access)
keywords: vbaac10.chm13792
f1_keywords:
- vbaac10.chm13792
ms.prod: access
api_name:
- Access.Report.Module
ms.assetid: e0cff3db-1697-7b8e-3934-7ead204052fb
ms.date: 06/08/2017
---


# Report.Module Property (Access)

You can use the  **Module** property to specify a report module. Read-only **Module** object.


## Syntax

 _expression_. **Module**

 _expression_ A variable that represents a **Report** object.


## Remarks

The  **Module** property also returns a reference to a specified **Module** object.

Use the  **Module** property to access the properties and methods of a **Module** object associated with a **Form** or **Report** object.

The setting of the  **HasModule** property of a form or report determines whether it has an associated module. If the **HasModule** property is **False**, the form or report does not have an associated module. When you refer to the **Module** property of that form or report while in design view, Microsoft Access creates the associated module and sets the **HasModule** property to **True**. If you refer to the **Module** property of a form or report at run-time and the object has its **HasModule** property set to **False**, an error will occur.

You could use this property with any of the properties and methods of the module object.


## Example

The following example uses the  **Module** property to insert the **Beep** method in a form's Open event.


```vb
Dim strFormOpenCode As String 
Dim mdl As Module 
 
Set mdl = Forms!MyForm.Module 
strFormOpenCode = "Sub Form_Open(Cancel As Integer)" _ 
 &; vbCrLf &; "Beep" &; vbCrLf &; "End Sub" 
 With mdl 
 .InsertText strFormOpenCode 
 End With
```


## See also


#### Concepts


[Report Object](report-object-access.md)

