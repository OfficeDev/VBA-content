---
title: Options.DefaultFilePath Property (Word)
keywords: vbawd10.chm162988097
f1_keywords:
- vbawd10.chm162988097
ms.prod: word
api_name:
- Word.Options.DefaultFilePath
ms.assetid: 39c90157-1824-55ee-c7e1-3687f132131f
ms.date: 06/08/2017
---


# Options.DefaultFilePath Property (Word)

Returns or sets default folders for items such as documents, templates, and graphics. Read/write  **String** .


## Syntax

 _expression_ . **DefaultFilePath**( **_Path_** )

 _expression_ Required. A variable that represents an **[Options](options-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Path_|Required| **WdDefaultFilePath**|The default folder to set.|

## Remarks

 You can use an empty string ("") to remove the setting from the Windows registry. The new setting takes effect immediately.


## Example

This example sets the default folder for Word documents.


```
Options.DefaultFilePath(wdDocumentsPath) = "C:\Documents"
```

This example returns the current default path for user templates (corresponds to the default path setting on the  **File Locations** tab in the **Options** dialog box).




```vb
Dim strPath As String 
 
strPath = Options.DefaultFilePath(wdUserTemplatesPath)
```


## See also


#### Concepts


[Options Object](options-object-word.md)

