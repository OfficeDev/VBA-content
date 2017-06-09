---
title: Application.Modules Property (Access)
keywords: vbaac10.chm12565
f1_keywords:
- vbaac10.chm12565
ms.prod: access
api_name:
- Access.Application.Modules
ms.assetid: eb99e25f-9a31-82cd-1b61-41c8a227b859
ms.date: 06/08/2017
---


# Application.Modules Property (Access)

You can use the  **Modules** property to access the **[Modules](modules-object-access.md)** collection and its related properties. Read-only **Modules** object.


## Syntax

 _expression_. **Modules**

 _expression_ A variable that represents an **Application** object.


## Remarks

Use the properties of the  **Modules** collection in Visual Basic to refer to all open standard modules and class modules.


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


[Application Object](application-object-access.md)

