---
title: Dialog.DefaultTab Property (Word)
keywords: vbawd10.chm163085570
f1_keywords:
- vbawd10.chm163085570
ms.prod: word
api_name:
- Word.Dialog.DefaultTab
ms.assetid: 22de708e-fb23-b27a-00f0-dc43787c7eaf
ms.date: 06/08/2017
---


# Dialog.DefaultTab Property (Word)

Returns or sets the active tab when the specified dialog box is displayed. Read/write  **WdWordDialogTab** .


## Syntax

 _expression_ . **DefaultTab**

 _expression_ Required. A variable that represents a **[Dialog](dialog-object-word.md)** object.


## Example

This example displays the  **Page Setup** dialog box with the **Paper Source** tab selected.


```vb
With Dialogs(wdDialogFilePageSetup) 
 .DefaultTab = wdDialogFilePageSetupTabPaperSource 
 .Show 
End With
```


## See also


#### Concepts


[Dialog Object](dialog-object-word.md)

