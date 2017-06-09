---
title: Version.SavedBy Property (Word)
keywords: vbawd10.chm162792427
f1_keywords:
- vbawd10.chm162792427
ms.prod: word
api_name:
- Word.Version.SavedBy
ms.assetid: 4e92d644-48e2-8dd7-ffef-9b626e4ca908
ms.date: 06/08/2017
---


# Version.SavedBy Property (Word)

Returns the name of the user who saved the specified version of the document. Read-only  **String** .


## Syntax

 _expression_ . **SavedBy**

 _expression_ An expression that returns a **[Version](version-object-word.md)** object.


## Example

This example displays the name of the user who saved the first version of the active document.


```vb
If ActiveDocument.Versions.Count >= 1 Then 
 MsgBox ActiveDocument.Versions(1).SavedBy 
End If
```

This example saves a version of the document with a comment and then displays the user name.




```vb
ActiveDocument.Versions.Save Comment:="Added client information" 
last = ActiveDocument.Versions.Count 
MsgBox ActiveDocument.Versions(last).SavedBy
```


## See also


#### Concepts


[Version Object](version-object-word.md)

