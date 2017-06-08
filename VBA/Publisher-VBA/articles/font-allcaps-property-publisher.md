---
title: Font.AllCaps Property (Publisher)
keywords: vbapb10.chm5373959
f1_keywords:
- vbapb10.chm5373959
ms.prod: publisher
api_name:
- Publisher.Font.AllCaps
ms.assetid: e8394f91-de31-0075-51ac-8a372023f0ce
ms.date: 06/08/2017
---


# Font.AllCaps Property (Publisher)

Returns or sets  **msoTrue** if the font is formatted as all capital letters, or returns one of the other **MsoTriState** constants if it is not. Read/write.


## Syntax

 _expression_. **AllCaps**

 _expression_A variable that represents a  **Font** object.


### Return Value

MsoTriState


## Remarks

Setting the  **AllCaps** property to **msoTrue** sets the **SmallCaps** property to **msoFalse**, and vice versa.

The  **AllCaps** property value can be one of the ** [MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** constants declared in the Microsoft Office type library.


## Example

This example checks the selected text in the active document for text formatted as all capital letters. For this example to work, there must be be an active publication with text selected.


```vb
Public Sub Caps() 
 
 If Publisher.ActiveDocument.Selection _ 
 .TextRange.Font.AllCaps = msoTrue Then 
 MsgBox "Text is all caps." 
 Else 
 MsgBox "Text is not all caps." 
 End If 
 
End Sub
```

This example formats the selected text as all capital letters. For this code to execute properly, an active document must exist with selected text.




```vb
Public Sub MakeCaps() 
 
 If Publisher.ActiveDocument.Selection.TextRange _ 
 .Font.AllCaps = msoFalse Then 
 Selection.TextRange.Font.AllCaps = msoTrue 
 Else 
 MsgBox "You need to select some text or it is already all caps." 
 End If 
 
End Sub
```


