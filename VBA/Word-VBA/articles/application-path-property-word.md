---
title: Application.Path Property (Word)
keywords: vbawd10.chm158335057
f1_keywords:
- vbawd10.chm158335057
ms.prod: word
api_name:
- Word.Application.Path
ms.assetid: 224b4c66-f49c-55f1-8b6b-74f5ed979a3d
ms.date: 06/08/2017
---


# Application.Path Property (Word)

Returns the disk or Web path to the specified object. Read-only  **String** .


## Syntax

 _expression_ . **Path**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


## Remarks

The path doesn't include a trailing character — for example, "C:\MSOffice" or "http://MyServer". Use the  **[PathSeparator](application-pathseparator-property-word.md)** property to add the character that separates folders and drive letters. Use the **[Name](document-name-property-word.md)** property of the **[Document](document-object-word.md)** object to return the file name without the path and use the **[FullName](document-fullname-property-word.md)** property to return the file name and the path together.


 **Note**  You can use the  **PathSeparator** property to build Web addresses even though they contain forward slashes (/) and the **PathSeparator** property defaults to a backslash (\).


## Example

This example displays the path and file name of the active document.


```vb
MsgBox ActiveDocument.Path &; Application.PathSeparator &; _ 
 ActiveDocument.Name
```

This example changes the current folder to the path of the template attached to the active document.




```
ChDir ActiveDocument.AttachedTemplate.Path
```

This example displays the path of the first add-in in the AddIns collection.




```vb
If AddIns.Count >= 1 Then MsgBox AddIns(1).Path
```


## See also


#### Concepts


[Application Object](application-object-word.md)

