---
title: Options.CreateBackup Property (Word)
keywords: vbawd10.chm162988073
f1_keywords:
- vbawd10.chm162988073
ms.prod: word
api_name:
- Word.Options.CreateBackup
ms.assetid: 02933ae3-1c3b-8309-d496-09c44d28a616
ms.date: 06/08/2017
---


# Options.CreateBackup Property (Word)

 **True** if Word creates a backup copy each time a document is saved. Read/write **Boolean** .


## Syntax

 _expression_ . **CreateBackup**

 _expression_ A variable that represents a **[Options](options-object-word.md)** object.


## Remarks

The  **CreateBackup** and **AllowFastSave** properties cannot be set to **True** concurrently.


## Example

This example sets Word to automatically create a backup copy, and then it saves the active document.


```vb
Options.CreateBackup = True 
ActiveDocument.Save
```

This example returns the current status of the  **Always create backup copy** option on the **Save** tab in the **Options** dialog box.




```vb
Dim blnBackup As Boolean 
 
blnBackup = Options.CreateBackup
```


## See also


#### Concepts


[Options Object](options-object-word.md)

