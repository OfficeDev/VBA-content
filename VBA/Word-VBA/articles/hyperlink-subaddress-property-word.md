---
title: Hyperlink.SubAddress Property (Word)
keywords: vbawd10.chm161285197
f1_keywords:
- vbawd10.chm161285197
ms.prod: word
api_name:
- Word.Hyperlink.SubAddress
ms.assetid: 9dff8453-c7e5-fd1a-89f8-869f762b0bdc
ms.date: 06/08/2017
---


# Hyperlink.SubAddress Property (Word)

Returns or sets a named location in the destination of the specified hyperlink. Read/write  **String** .


## Syntax

 _expression_ . **SubAddress**

 _expression_ An expression that returns a **[Hyperlink](hyperlink-object-word.md)** object.


## Remarks

The named location can be a bookmark in a Microsoft Word document, a named cell or cell reference in a Microsoft Excel worksheet, a named object in a Microsoft Access database, or a slide number in a Microsoft PowerPoint presentation.


## Example

This example displays the subaddress of the selected hyperlink.


```vb
If Selection.Range.Hyperlinks.Count >= 1 Then 
 MsgBox Selection.Range.Hyperlinks(1).SubAddress 
End If
```

This example adds a hyperlink to the selection in the active document, sets the hyperlink destination and subaddress, and then displays them in a message box.




```vb
Set SCut = ActiveDocument.Hyperlinks.Add( _ 
 Anchor:= Selection.Range, _ 
 Address:="C:\My Documents\Other.doc", SubAddress:= "temp") 
MsgBox "The hyperlink goes to " &; SCut.SubAddress
```


## See also


#### Concepts


[Hyperlink Object](hyperlink-object-word.md)

