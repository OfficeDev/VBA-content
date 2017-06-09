---
title: Shape.Title Property (Word)
keywords: vbawd10.chm161480862
f1_keywords:
- vbawd10.chm161480862
ms.prod: word
api_name:
- Word.Shape.Title
ms.assetid: bb7c0810-8148-6123-033d-1d6de529dffa
ms.date: 06/08/2017
---


# Shape.Title Property (Word)

Returns or sets a  **String** that contains a title for the specified shape. Read/write.


## Syntax

 _expression_ . **Title**

 _expression_ A variable that represents a **[Shape](shape-object-word.md)** object.


## Remarks

Use the  **Title** property to provide an alternative text title for a shape. This property adds title text to the **Title** text box on the **Alt Text** pane of the **Format Shape** dialog in Word.


 **Note**  Web browsers display alternative text while tables are loading or if they are missing. Web search engines use the alternative text to help find Web pages. Alternative text is also used to assist disabilities.


## Example

The following code example adds an alternative text title to the second shape in the active document.


```vb
ActiveDocument.Shapes(2).Title = "Shape 2."
```


## See also


#### Concepts


[Shape Object](shape-object-word.md)

