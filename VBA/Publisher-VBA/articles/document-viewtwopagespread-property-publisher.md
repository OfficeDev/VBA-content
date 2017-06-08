---
title: Document.ViewTwoPageSpread Property (Publisher)
keywords: vbapb10.chm196665
f1_keywords:
- vbapb10.chm196665
ms.prod: publisher
api_name:
- Publisher.Document.ViewTwoPageSpread
ms.assetid: b5e851ff-d5fc-a98d-02b3-7e14c1b957dc
ms.date: 06/08/2017
---


# Document.ViewTwoPageSpread Property (Publisher)

Returns  **True** if the specified publication should be viewed as a two-page spread. Read/write **Boolean**.


## Syntax

 _expression_. **ViewTwoPageSpread**

 _expression_A variable that represents a  **Document** object.


### Return Value

Boolean


## Example

This example opens a message box and displays if the current publication should be viewed in the in the two page spread mode.


```vb
Sub ViewTwoPage() 
 
 MsgBox "View Two Page Spread = " &; _ 
 Application.ActiveDocument.ViewTwoPageSpread 
 
End Sub
```


