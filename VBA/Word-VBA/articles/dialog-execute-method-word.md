---
title: Dialog.Execute Method (Word)
keywords: vbawd10.chm163085569
f1_keywords:
- vbawd10.chm163085569
ms.prod: word
api_name:
- Word.Dialog.Execute
ms.assetid: 7f7dce3a-40ef-988c-f5ea-06a25c0ccc4b
ms.date: 06/08/2017
---


# Dialog.Execute Method (Word)

Applies the current settings of a Microsoft Word dialog box.


## Syntax

 _expression_ . **Execute**

 _expression_ Required. A variable that represents a **[Dialog](dialog-object-word.md)** object.


## Example

The following example enables the  **Keep with next** check box on the **Line and Page Breaks** tab in the **Paragraph** dialog box.


```vb
With Dialogs(wdDialogFormatParagraph) 
 .KeepWithNext = 1 
 .Execute 
End With
```


## See also


#### Concepts


[Dialog Object](dialog-object-word.md)

