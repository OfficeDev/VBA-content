---
title: Style.BaseStyle Property (Word)
keywords: vbawd10.chm153878529
f1_keywords:
- vbawd10.chm153878529
ms.prod: word
api_name:
- Word.Style.BaseStyle
ms.assetid: d055a10a-66c4-7b50-923c-ab60fde0efa9
ms.date: 06/08/2017
---


# Style.BaseStyle Property (Word)

Returns or sets an existing style on which you can base the formatting of another style. Read/write  **Variant** .


## Syntax

 _expression_ . **BaseStyle**

 _expression_ A variable that represents a **[Style](style-object-word.md)** object.


## Remarks

To set the  **BaseStyle** property, specify either the local name of the base style, an integer or a **wdBuiltinStyle** constant, or an object that represents the base style. For a list of the **wdBuiltinStyle** constants, see the **Style** property for the object that you want to set.


## Example

This example creates a new document and then adds a new paragraph style named "myHeading." It assigns Heading 1 as the base style for the new style. A left indent of 1 inch (72 points) is then specified for the new style.


```vb
Dim docNew As Document 
Dim styleNew As Style 
 
Set docNew = Documents.Add 
Set styleNew = docNew.Styles.Add("NewHeading1") 
With styleNew 
 .BaseStyle = docNew.Styles(wdStyleHeading1) 
 .ParagraphFormat.LeftIndent = 72 
End With
```

This example returns the base style that's used for the Body Text paragraph style.




```vb
Dim styleBase As Style 
 
styleBase = ActiveDocument.Styles(wdStyleBodyText).BaseStyle 
MsgBox styleBase
```


## See also


#### Concepts


[Style Object](style-object-word.md)

