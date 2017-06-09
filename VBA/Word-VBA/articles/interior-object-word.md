---
title: Interior Object (Word)
keywords: vbawd10.chm43
f1_keywords:
- vbawd10.chm43
ms.prod: word
api_name:
- Word.Interior
ms.assetid: 6fc3e311-a7c9-bfa9-7459-9cea177b08e5
ms.date: 06/08/2017
---


# Interior Object (Word)

Represents the interior of an object.


## Example

The following example enables up and down bars, and then sets the interior color of the up bars to green, for the first chart group of the first chart in the active document. Use the  **[UpBars.Interior](http://msdn.microsoft.com/library/89584c60-be0f-45e8-4d45-86c6c7806c44%28Office.15%29.aspx)** property to return the **Interior** object.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.ChartGroups(1) 
 .HasUpDownBars = True 
 .UpBars.Interior.ColorIndex = 4 
 End With 
 End If 
End With
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


