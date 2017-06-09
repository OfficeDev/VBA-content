---
title: DefaultWebOptions.CheckIfWordIsDefaultHTMLEditor Property (Word)
keywords: vbawd10.chm165871624
f1_keywords:
- vbawd10.chm165871624
ms.prod: word
api_name:
- Word.DefaultWebOptions.CheckIfWordIsDefaultHTMLEditor
ms.assetid: 9d3fbbe1-3a21-64eb-1266-ce22b2332e61
ms.date: 06/08/2017
---


# DefaultWebOptions.CheckIfWordIsDefaultHTMLEditor Property (Word)

 **True** if Microsoft Word checks to see whether it is the default HTML editor when you start Word. Read/write **Boolean** .


## Syntax

 _expression_ . **CheckIfWordIsDefaultHTMLEditor**

 _expression_ A variable that represents a **[DefaultWebOptions](defaultweboptions-object-word.md)** object.


## Remarks

The  **CheckIfWordIsDefaultHTMLEditor** property returns **False** if Word does not perform this check. The default value is **True** .

This property is used only if the Web browser you are using supports HTML editing and HTML editors. To use a different HTML editor, you must set this property to  **False** and then register the editor as the default system HTML editor.


## Example

This example sets Microsoft Word to check to see whether it is the default HTML editor.


```vb
Application.DefaultWebOptions _ 
 .CheckIfWordIsDefaultHTMLEditor = True
```


## See also


#### Concepts


[DefaultWebOptions Object](defaultweboptions-object-word.md)

