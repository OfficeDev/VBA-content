---
title: Series.Paste Method (Word)
keywords: vbawd10.chm123732179
f1_keywords:
- vbawd10.chm123732179
ms.prod: word
api_name:
- Word.Series.Paste
ms.assetid: cef0e06e-fc4d-b63f-aea6-4cd325c3e0b9
ms.date: 06/08/2017
---


# Series.Paste Method (Word)

Pastes a picture from the Clipboard as the marker on the selected series.


## Syntax

 _expression_ . **Paste**

 _expression_ A variable that represents a **[Series](series-object-word.md)** object.


## Remarks

You can use this method on column, bar, line, or radar charts, and it sets the  **[MarkerStyle](series-markerstyle-property-word.md)** property to **xlMarkerStylePicture** .


## Example

The following example pastes a picture from the Clipboard into series one for the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).Paste 
 End If 
End With 

```


## See also


#### Concepts


[Series Object](series-object-word.md)

