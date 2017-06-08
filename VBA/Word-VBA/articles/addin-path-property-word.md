---
title: AddIn.Path Property (Word)
keywords: vbawd10.chm159252483
f1_keywords:
- vbawd10.chm159252483
ms.prod: word
api_name:
- Word.AddIn.Path
ms.assetid: 0c9150fe-a57f-85d5-275b-a45916c35f76
ms.date: 06/08/2017
---


# AddIn.Path Property (Word)

Returns the location of an installed add-in. Read-only  **String** .


## Syntax

 _expression_ . **Path**

 _expression_ An expression that returns an **[AddIn](addin-object-word.md)** object.


## Remarks

The path doesn't include a trailing character— for example, "C:\MSOffice" or "http://MyServer". Use the  **PathSeparator** property to add the character that separates folders and drive letters. Use the **Name** property to return the file name without the path and use the **FullName** property to return the file name and the path together.


 **Note**  You can use the  **PathSeparator** property to build Web addresses even though they contain forward slashes (/) and the **PathSeparator** property defaults to a backslash (\).


## Example

This example displays the path of the first add-in in the  **AddIns** collection.


```vb
If AddIns.Count >= 1 Then MsgBox AddIns(1).Path
```


## See also


#### Concepts


[AddIn Object](addin-object-word.md)

