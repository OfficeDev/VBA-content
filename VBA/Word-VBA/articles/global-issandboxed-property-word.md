---
title: Global.IsSandboxed Property (Word)
keywords: vbawd10.chm163119220
f1_keywords:
- vbawd10.chm163119220
ms.prod: word
api_name:
- Word.Global.IsSandboxed
ms.assetid: 12bef36b-7ec6-5b43-f8b8-dbb5dacef868
ms.date: 06/08/2017
---


# Global.IsSandboxed Property (Word)

 **True** if the application window is a protected view window. Read-only.


## Syntax

 _expression_ . **IsSandboxed**

 _expression_ An expression that returns a **[Global](global-object-word.md)** object.


## Example

The following code example displays whether or not the document referenced by  _doc_ is in a protected view window.


```
If doc.Application.IsSandboxed Then 
 MsgBox "The document " &; _ 
 """" &; doc.Name &; """" &; _ 
 " is in a protected view window." 
Else 
 MsgBox "The document " &; _ 
 """" &; doc.Name &; """" &; _ 
 " is not in a protected view window." 
End If
```


## See also


#### Concepts


[Global Object](global-object-word.md)

