---
title: View.DisplayPageBoundaries Property (Word)
keywords: vbawd10.chm161808416
f1_keywords:
- vbawd10.chm161808416
ms.prod: word
api_name:
- Word.View.DisplayPageBoundaries
ms.assetid: 67b91767-c9aa-6d2e-d99b-258a79777c25
ms.date: 06/08/2017
---


# View.DisplayPageBoundaries Property (Word)

 **True** to display the top and bottom margins (white space) and the gray area (gray space) between pages in a document. Read/write **Boolean** .


## Syntax

 _expression_ . **DisplayPageBoundaries**

 _expression_ An expression that returns a **[View](view-object-word.md)** object.


## Remarks

 **False** to hide from view the white and gray space so that the pages flow together as one long page. The default value is **True** .

This feature is only available in the Print Layout view and only affects the gray space on the top and bottom of a page, not the left and right sides of a page. This setting affects the document in the in the specified window. When the document is saved, the state of this setting is saved with it.


## Example

This example changes the current view to Print Layout and suppresses the white and gray space between document pages.


```vb
Sub WhiteSpace() 
 With ActiveWindow.View 
 .Type = wdPrintView 
 .DisplayPageBoundaries = False 
 End With 
End Sub
```


## See also


#### Concepts


[View Object](view-object-word.md)

