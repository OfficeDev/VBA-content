---
title: DefaultWebOptions.UpdateLinksOnSave Property (Word)
keywords: vbawd10.chm165871621
f1_keywords:
- vbawd10.chm165871621
ms.prod: word
api_name:
- Word.DefaultWebOptions.UpdateLinksOnSave
ms.assetid: f926c078-ae86-fa73-8201-568e3cba2306
ms.date: 06/08/2017
---


# DefaultWebOptions.UpdateLinksOnSave Property (Word)

 **True** if hyperlinks and paths to all supporting files are automatically updated before you save the document as a Web page. Read/write **Boolean** .


## Syntax

 _expression_ . **UpdateLinksOnSave**

 _expression_ An expression that returns a **[DefaultWebOptions](defaultweboptions-object-word.md)** object.


## Remarks

The  **UpdateLinksOnSave** property ensures that the links are up-to-date at the time the document is saved. The default value for the **UpdateLinksOnSave** property is **True** .

A value of  **False** indicates that the links are not updated. You should set this property to **False** if the location where the document is saved is different from the final location on the Web server and the supporting files are not available at the first location.


## Example

This example specifies that links are not updated before the document is saved.


```vb
Application.DefaultWebOptions.UpdateLinksOnSave = False
```


## See also


#### Concepts


[DefaultWebOptions Object](defaultweboptions-object-word.md)

