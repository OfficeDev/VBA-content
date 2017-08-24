---
title: PictureFormat.TransparencyColor Property (Publisher)
keywords: vbapb10.chm3604743
f1_keywords:
- vbapb10.chm3604743
ms.prod: publisher
api_name:
- Publisher.PictureFormat.TransparencyColor
ms.assetid: 908d2e21-3e2a-b75b-a82d-454686b7ecb8
ms.date: 06/08/2017
---


# PictureFormat.TransparencyColor Property (Publisher)

Returns or sets an  **MsoRGBType** constant that represents the transparency color. Read/write.


## Syntax

 _expression_. **TransparencyColor**

 _expression_A variable that represents a  **PictureFormat** object.


### Return Value

MsoRGBType


## Example

This example creates a picture on the first page and sets the transparency color to black.


```vb
Sub SetTransparentColor() 
 With ActiveDocument.Pages(1).Shapes.AddPicture( _ 
 FileName:="C:\My Pictures\Sample.gif", LinkToFile:=msoFalse, _ 
 SaveWithDocument:=msoTrue, Left:=36, Top:=36) 
 .PictureFormat.TransparencyColor = RGB(Red:=255, Green:=255, Blue:=255) 
 End With 
End Sub
```


