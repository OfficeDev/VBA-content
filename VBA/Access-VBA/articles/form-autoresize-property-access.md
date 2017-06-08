---
title: Form.AutoResize Property (Access)
keywords: vbaac10.chm13367
f1_keywords:
- vbaac10.chm13367
ms.prod: access
api_name:
- Access.Form.AutoResize
ms.assetid: 5ae98bc8-fa33-7e4b-31c8-ba22aa026a45
ms.date: 06/08/2017
---


# Form.AutoResize Property (Access)

Returns or sets a  **Boolean** indicating whether a Form window opens automatically sized to display complete records. Read/write.


## Syntax

 _expression_. **AutoResize**

 _expression_ A variable that represents a **Form** object.


## Remarks

The  **AutoResize** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Yes|**True**|(Default) The Form window is automatically sized to display a complete record.|
|No|**False**|When opened, the Form window has the last saved size. To save a window's size, open the form, size the window, save the form by clicking  **Save** on the **File** menu, and close the form or report. When you next open the form or report, it will be the saved window size.|
This property can be set only in Design view.

The Form window resizes only if opened in Form view. If you open the form first in Design view or Datasheet view and then change to Form view, the Form window won't resize.

If you make any changes in Design view to a form whose  **AutoResize** property is set to No and whose **AutoCenter** property is set to Yes, switch to Form view before saving the form. If you don't, Microsoft Access clips the form on the right and bottom edges the next time you open the form.

If the  **AutoCenter** property is set to No, a Form window opens with its upper-left corner in the same location as when it was closed.


## See also


#### Concepts


[Form Object](form-object-access.md)

