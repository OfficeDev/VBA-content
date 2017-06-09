---
title: TextStyle.Description Property (Publisher)
keywords: vbapb10.chm5963779
f1_keywords:
- vbapb10.chm5963779
ms.prod: publisher
api_name:
- Publisher.TextStyle.Description
ms.assetid: 278d647e-c4bc-218c-417d-b01caf2d98a3
ms.date: 06/08/2017
---


# TextStyle.Description Property (Publisher)

Returns a  **String** that represents the description of the specified style. For example, a typical description for the Normal style might be "(Default) Times New Roman, (Asian) MS Mincho, 10 pt, Main (Black), Kerning 14 pt, Left, Line spacing 1 sp." Read-only.


## Syntax

 _expression_. **Description**

 _expression_A variable that represents a  **TextStyle** object.


## Example

This example displays the description for the Normal style.


```vb
Sub ShowStyleDescription() 
 MsgBox "The Normal style has the following formatting attributes: " &; _ 
 vbLf &; ActiveDocument.TextStyles("Normal").Description 
End Sub
```


