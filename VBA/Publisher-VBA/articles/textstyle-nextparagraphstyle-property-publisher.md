---
title: TextStyle.NextParagraphStyle Property (Publisher)
keywords: vbapb10.chm5963784
f1_keywords:
- vbapb10.chm5963784
ms.prod: publisher
api_name:
- Publisher.TextStyle.NextParagraphStyle
ms.assetid: 2b31b883-c26d-3be8-7145-f8e3cf1ba5cc
ms.date: 06/08/2017
---


# TextStyle.NextParagraphStyle Property (Publisher)

Returns or sets a  **String** that represents the paragraph style that follows the specified text style when a user presses ENTER. Read/write.


## Syntax

 _expression_. **NextParagraphStyle**

 _expression_A variable that represents a  **TextStyle** object.


### Return Value

String


## Example

This example creates a new text style and specifies that the text style following the new text style is the Normal style.


```vb
Sub CreateNewTextStyle() 
 Dim styNew As TextStyle 
 Dim fntStyle As Font 
 
 Set styNew = ActiveDocument.TextStyles.Add(StyleName:="Heading 1") 
 Set fntStyle = styNew.Font 
 
 With fntStyle 
 .Name = "Tahoma" 
 .Bold = msoTrue 
 .Size = 15 
 End With 
 
 With styNew 
 .Font = fntStyle 
 .NextParagraphStyle = "Normal" 
 End With 
End Sub
```


