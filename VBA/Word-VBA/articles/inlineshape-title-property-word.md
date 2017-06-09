---
title: InlineShape.Title Property (Word)
keywords: vbawd10.chm162005150
f1_keywords:
- vbawd10.chm162005150
ms.prod: word
api_name:
- Word.InlineShape.Title
ms.assetid: 85a28df8-f579-79b9-60b9-30624a64dae7
ms.date: 06/08/2017
---


# InlineShape.Title Property (Word)

Returns or sets a  **String** that contains a title for the specified inline shape. Read/write.


## Syntax

 _expression_ . **Title**

 _expression_ A variable that represents an **[InlineShape](inlineshape-object-word.md)** object.


## Remarks

Use the  **Title** property to provide an alternative text title for an inline shape. This property adds title text to the **Title** text box on the **Alt Text** pane of the **Format Shape** dialog in Word.


 **Note**  Web browsers display alternative text while tables are loading or if they are missing. Web search engines use the alternative text to help find Web pages. Alternative text is also used to assist disabilities.


## Example

The following code example adds an alternative text title to the first inline shape in the active document. 


```vb
ActiveDocument.InlineShapes(1).Title = "Desktop screenshot."
```


## See also


#### Concepts


[InlineShape Object](inlineshape-object-word.md)

