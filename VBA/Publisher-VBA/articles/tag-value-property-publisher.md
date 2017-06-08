---
title: Tag.Value Property (Publisher)
keywords: vbapb10.chm4718596
f1_keywords:
- vbapb10.chm4718596
ms.prod: publisher
api_name:
- Publisher.Tag.Value
ms.assetid: dee3b69b-ae5b-df13-561e-84105057979a
ms.date: 06/08/2017
---


# Tag.Value Property (Publisher)

Returns or sets a  **Variant** that represents the value of a tag of a shape, page, or publication. Read/write.


## Syntax

 _expression_. **Value**

 _expression_A variable that represents a  **Tag** object.


## Example

This example creates a new tag for the active publication and then displays the value of the tag.


```vb
Sub CreatePublicationTag() 
 With ActiveDocument 
 .Tags.Add Name:="ActivePub", Value:="This is the active publication." 
 MsgBox .Tags(1).Value 
 End With 
End Sub
```


