---
title: Font.Grow Method (Publisher)
keywords: vbapb10.chm5373990
f1_keywords:
- vbapb10.chm5373990
ms.prod: publisher
api_name:
- Publisher.Font.Grow
ms.assetid: 41d48db2-4a0d-6efc-80c5-c6f035e9e6ff
ms.date: 06/08/2017
---


# Font.Grow Method (Publisher)

Increases the font size to the next available size.


## Syntax

 _expression_. **Grow**

 _expression_A variable that represents a  **Font** object.


## Remarks

If the selection or range contains more than one font size, each size is increased to the next available setting.


## Example

This example increases the font size of the fourth word in a new textbox.


```vb
Sub GrowFont() 
 Dim shpText As Shape 
 Dim intResponse As Integer 
 
 Set shpText = ActiveDocument.Pages(1).Shapes.AddTextbox( _ 
 Orientation:=pbTextOrientationHorizontal, Left:=100, _ 
 Top:=100, Width:=200, Height:=100) 
 
 With shpText.TextFrame.TextRange 
 .Text = "This is a test of the Grow method." 
 Do Until intResponse = vbNo 
 intResponse = MsgBox("Do you want to increase the " &; _ 
 "size of the font?", vbYesNo) 
 If intResponse = vbYes Then 
 .Words(4).Font.Grow 
 End If 
 Loop 
 End With 
End Sub
```

This example increases the font size of the selected text.




```vb
Sub IncreaseFontSizeOfSelectedText() 
 If Selection.Type = pbSelectionText Then 
 Selection.TextRange.Font.Grow 
 Else 
 MsgBox "You need to select some text." 
 End If 
End Sub
```


