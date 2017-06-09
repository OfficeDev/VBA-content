---
title: Form.AutoCenter Property (Access)
keywords: vbaac10.chm13368
f1_keywords:
- vbaac10.chm13368
ms.prod: access
api_name:
- Access.Form.AutoCenter
ms.assetid: a60f8783-5a25-42b5-da99-c5e2925fd6ea
ms.date: 06/08/2017
---


# Form.AutoCenter Property (Access)

Returns or sets a  **Boolean** indicating whether a form will be centered automatically in the application window when the form is opened. Read/write.


## Syntax

 _expression_. **AutoCenter**

 _expression_ A variable that represents a **Form** object.


## Remarks

The  **AutoCenter** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Yes|**True**|The form will be centered automatically on opening.|
|No|**False**|(Default) The form upper-left corner will be in the same location as when the form was last saved.|
You can set this property only in Design view.

Depending on the size and placement of the application window, forms can appear off to one side of the application window, hiding part of the form or report. Centering the form automatically when it's opened makes it easier to view and use.

If you make any changes in Design view to a form whose  **AutoResize** property is set to No and whose **AutoCenter** property is set to Yes, switch to Form view before saving the form. If you don't, Microsoft Access clips the form on the right and bottom edges the next time you open the form.


## See also


#### Concepts


[Form Object](form-object-access.md)

