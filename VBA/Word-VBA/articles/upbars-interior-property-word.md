---
title: UpBars.Interior Property (Word)
keywords: vbawd10.chm180945025
f1_keywords:
- vbawd10.chm180945025
ms.prod: word
api_name:
- Word.UpBars.Interior
ms.assetid: 2ea3eef1-4602-c81a-852b-e6e4f11d2065
ms.date: 06/08/2017
---


# UpBars.Interior Property (Word)

Returns the interior of the object. Read-only  **[Interior](interior-object-word.md)** .


## Syntax

 _expression_ . **Interior**

 _expression_ A variable that represents an **[UpBars](upbars-object-word.md)** object.


## Example

The following example enables up and down bars, then sets the interior color of the down bars to red and the up bars to green, for the first chart group of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.ChartGroups(1) 
 .HasUpDownBars = True 
 .DownBars.Interior.ColorIndex = 3 
 .UpBars.Interior.ColorIndex = 4 
 End With 
 End If 
End With
```


## See also


#### Concepts


[UpBars Object](upbars-object-word.md)

