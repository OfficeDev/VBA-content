---
title: DefaultWebOptions.CheckIfOfficeIsHTMLEditor Property (Word)
keywords: vbawd10.chm165871623
f1_keywords:
- vbawd10.chm165871623
ms.prod: word
api_name:
- Word.DefaultWebOptions.CheckIfOfficeIsHTMLEditor
ms.assetid: 5475aaff-70df-cb52-7bf7-d58f8c27fffb
ms.date: 06/08/2017
---


# DefaultWebOptions.CheckIfOfficeIsHTMLEditor Property (Word)

 **True** if Microsoft Word checks to see whether an Office application is the default HTML editor when you start Word. Read/write **Boolean** .


## Syntax

 _expression_ . **CheckIfOfficeIsHTMLEditor**

 _expression_ A variable that represents a **[DefaultWebOptions](defaultweboptions-object-word.md)** object.


## Remarks

The  **CheckIfOfficeIsHTMLEditor** property returns **False** if Word does not perform this check. The default value is **True** .

 This property is used only if the Web browser you are using supports HTML editing and HTML editors. To use a different HTML editor, you must set this property to **False** and then register the editor as the default system HTML editor.


## Example

This example causes Microsoft Word not to check to see whether an Office application is the default HTML editor.


```vb
Application.DefaultWebOptions _ 
 .CheckIfOfficeIsHTMLEditor = False
```


## See also


#### Concepts


[DefaultWebOptions Object](defaultweboptions-object-word.md)

