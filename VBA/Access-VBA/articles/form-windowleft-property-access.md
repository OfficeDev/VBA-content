---
title: Form.WindowLeft Property (Access)
keywords: vbaac10.chm13517
f1_keywords:
- vbaac10.chm13517
ms.prod: access
api_name:
- Access.Form.WindowLeft
ms.assetid: f9e90b5e-6008-675d-9168-6dd932559b6d
ms.date: 06/08/2017
---


# Form.WindowLeft Property (Access)

Returns an  **Integer** indicating the screen position in twips of the left edge of a form relative to the left edge of the Microsoft Access window. Read-only.


## Syntax

 _expression_. **WindowLeft**

 _expression_ A variable that represents a **Form** object.


## Remarks

Use the  **Move** method to change the position of a form.


## Example

The following example returns the screen position of the top and left edges of the first form in the current project.


```vb
With Forms(0) 
 
 MsgBox "The form is " &; .WindowLeft _ 
 &; " twips from the left edge of the Access window and " _ 
 &; .WindowTop _ 
 &; " twips from the top edge of the Access window." 
 
End With 

```


## See also


#### Concepts


[Form Object](form-object-access.md)

