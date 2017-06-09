---
title: Form.Pages Property (Access)
keywords: vbaac10.chm13411
f1_keywords:
- vbaac10.chm13411
ms.prod: access
api_name:
- Access.Form.Pages
ms.assetid: 9494fb79-d080-e2cb-6b55-8194ecd81e9b
ms.date: 06/08/2017
---


# Form.Pages Property (Access)

You can use the  **Pages** property to return information needed to print page numbers in a form. Read/write **Integer**.


## Syntax

 _expression_. **Pages**

 _expression_ A variable that represents a **Form** object.


## Remarks

This property is only available in Print Preview or when printing.

To refer to the  **Pages** property in a macro or Visual Basic, the form or report must include a text box whose **ControlSource** property is set to an expression that uses **Pages**. For example, you can use the following expressions as the **ControlSource** property setting for a text box in a page footer.



|**This expression**|**Prints**|
|:-----|:-----|
|=Page|A page number (for example, 1, 2, 3) in the page footer.|
|="Page " &; Page &; " of " &; Pages|"Page  _n_ of _nn_ " (for example, Page 1 of 5, Page 2 of 5) in the page footer.|
|=Pages|The total number pages in the form (for example, 5).|

## See also


#### Concepts


[Form Object](form-object-access.md)

