---
title: Options.PrintDrawingObjects Property (Word)
keywords: vbawd10.chm162988070
f1_keywords:
- vbawd10.chm162988070
ms.prod: word
api_name:
- Word.Options.PrintDrawingObjects
ms.assetid: 366ddc26-1cb0-fe48-8d54-ff9d5d3492b4
ms.date: 06/08/2017
---


# Options.PrintDrawingObjects Property (Word)

 **True** if Microsoft Word prints drawing objects. Read/write **Boolean** .


## Syntax

 _expression_ . **PrintDrawingObjects**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Example

This example sets Word to print drawing objects, and then it prints the active document.


```vb
Options.PrintDrawingObjects = True 
ActiveDocument.PrintOut
```

This example returns the current status of the  **Drawing objects** option on the **Print** tab in the **Options** dialog box ( **Tools** menu).




```
temp = Options.PrintDrawingObjects
```


## See also


#### Concepts


[Options Object](options-object-word.md)

