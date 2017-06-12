---
title: Style.NameLocal Property (Word)
keywords: vbawd10.chm153878528
f1_keywords:
- vbawd10.chm153878528
ms.prod: word
api_name:
- Word.Style.NameLocal
ms.assetid: 49d5d7d7-65b5-2861-171b-3badfe055568
ms.date: 06/08/2017
---


# Style.NameLocal Property (Word)

Returns the name of a built-in style in the language of the user. Read/write  **String** .


## Syntax

 _expression_ . **NameLocal**

 _expression_ Required. A variable that represents a **[Style](style-object-word.md)** object.


## Remarks

Setting this property renames a user-defined style or adds an alias to a built-in style.


## Example

This example displays the style name (in the language of the user) applied to the selected paragraphs. If more than one style has been applied to the selection, the first style name is displayed.


```vb
MsgBox Selection.Paragraphs.Style.NameLocal
```

This example adds the name "MyH1" as the alias for the Heading 1 style in the active document.




```vb
ActiveDocument.Styles("Heading 1").NameLocal = "MyH1"
```

This example renames the style named "Test" to "Intro."




```vb
ActiveDocument.Styles("Test").NameLocal = "Intro"
```


## See also


#### Concepts


[Style Object](style-object-word.md)

