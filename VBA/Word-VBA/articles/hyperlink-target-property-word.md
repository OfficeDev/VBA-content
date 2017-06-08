---
title: Hyperlink.Target Property (Word)
keywords: vbawd10.chm161285109
f1_keywords:
- vbawd10.chm161285109
ms.prod: word
api_name:
- Word.Hyperlink.Target
ms.assetid: 2a36ec74-fcfd-9000-8229-dcd01b8f7757
ms.date: 06/08/2017
---


# Hyperlink.Target Property (Word)

Returns or sets the name of the frame or window in which to load the hyperlink. Read/write  **String** .


## Syntax

 _expression_ . **Target**

 _expression_ Required. A variable that represents a **[Hyperlink](hyperlink-object-word.md)** object.


## Example

This example sets the specified hyperlink to open in a new browser window.


```vb
ActiveDocument.Hyperlinks(1).Target = "_blank"
```

This example sets the specified hyperlink to open in the frame called "left."




```vb
ActiveDocument.Hyperlinks(1).Target = "left"
```


## See also


#### Concepts


[Hyperlink Object](hyperlink-object-word.md)

