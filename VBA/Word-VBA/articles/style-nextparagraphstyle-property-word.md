---
title: Style.NextParagraphStyle Property (Word)
keywords: vbawd10.chm153878533
f1_keywords:
- vbawd10.chm153878533
ms.prod: word
api_name:
- Word.Style.NextParagraphStyle
ms.assetid: f8326275-bb81-4a0e-f790-32b34ef71f78
ms.date: 06/08/2017
---


# Style.NextParagraphStyle Property (Word)

Returns or sets the style to be applied automatically to a new paragraph that is inserted after a paragraph formatted with the specified style. Read/write  **Variant** .


## Syntax

 _expression_ . **NextParagraphStyle**

 _expression_ An expression that returns a **[Style](style-object-word.md)** object.


## Remarks

You can set the  **NextParagraphStyle** property by using the local name of the style, an integer or a **WdBuiltinStyle** constant, or an object that represents the next style. For a list of the **WdBuiltinStyle** constants, see the **Style** property for the object that you want to set.


## Example

This example sets the Heading 1 style to be followed by the Heading 2 style in the active document.


```vb
ActiveDocument.Styles(wdStyleHeading1).NextParagraphStyle = _ 
 ActiveDocument.Styles(wdStyleHeading2)
```

This example creates a new document and adds a paragraph style named "MyStyle." The new style is based on the Normal style, is followed by the Heading 3 style, has a left indent of 1 inch (72 points), and is formatted as bold.




```vb
Set myDoc = Documents.Add 
Set myStyle = myDoc.Styles.Add(Name:= "MyStyle") 
 With myStyle 
 .BaseStyle = wdStyleNormal 
 .NextParagraphStyle = wdStyleHeading3 
 .ParagraphFormat.LeftIndent = 72 
 .Font.Bold = True 
 End With
```


## See also


#### Concepts


[Style Object](style-object-word.md)

