---
title: AddIn.Compiled Property (Word)
keywords: vbawd10.chm159252485
f1_keywords:
- vbawd10.chm159252485
ms.prod: word
api_name:
- Word.AddIn.Compiled
ms.assetid: 812402c2-8755-cb40-beb8-e46cfba3e0ea
ms.date: 06/08/2017
---


# AddIn.Compiled Property (Word)

 **True** if the specified add-in is a Word add-in library (WLL). **False** if the add-in is a template. Read-only **Boolean** .


## Syntax

 _expression_ . **Compiled**

 _expression_ A variable that represents a **[AddIn](addin-object-word.md)** object.


## Example

This example determines how many WLLs are currently loaded.


```vb
count = 0 
For Each aAddin in Addins 
 If aAddin.Compiled = True And aAddin.Installed = True Then 
 count = count + 1 
 End If 
Next aAddin 
MsgBox Str(count) &; " WLL's are loaded"
```

If the first add-in is a template, this example unloads the template and opens it.




```vb
If Addins(1).Compiled = False Then 
 Addins(1).Installed = False 
 Documents.Open FileName:=AddIns(1).Path _ 
 &; Application.PathSeparator _ 
 &; AddIns(1).Name 
End If
```


## See also


#### Concepts


[AddIn Object](addin-object-word.md)

