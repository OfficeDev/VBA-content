---
title: Hyperlink.ExtraInfoRequired Property (Word)
keywords: vbawd10.chm161285105
f1_keywords:
- vbawd10.chm161285105
ms.prod: word
api_name:
- Word.Hyperlink.ExtraInfoRequired
ms.assetid: 066a4dbf-f5ea-f708-cd57-f8e515a258d5
ms.date: 06/08/2017
---


# Hyperlink.ExtraInfoRequired Property (Word)

 **True** if extra information is required to resolve the specified hyperlink. Read-only **Boolean** .


## Syntax

 _expression_ . **ExtraInfoRequired**

 _expression_ A variable that represents a **[Hyperlink](hyperlink-object-word.md)** object.


## Remarks

You can specify extra information by using the ExtraInfo argument with the  **[Follow](hyperlink-follow-method-word.md)** or **[FollowHyperlink](document-followhyperlink-method-word.md)** method. For example, you can use ExtraInfo to specify the coordinates of an image map, the contents of a form, or a FAT file name.


## Example

This example inserts a hyperlink to www.msn.com and then follows the hyperlink if extra information isn't required.


```vb
Dim hypTemp As Hyperlink 
 
With Selection 
 .Collapse Direction:=wdCollapseEnd 
 .InsertAfter "MSN " 
 .Previous 
End With 
Set hypTemp = ActiveDocument.Hyperlinks.Add( _ 
 Address:="http://www.msn.com", _ 
 Anchor:=Selection.Range) 
If hypTemp.ExtraInfoRequired = False Then 
 hypTemp.Follow 
End If
```


## See also


#### Concepts


[Hyperlink Object](hyperlink-object-word.md)

