---
title: Trendline.InterceptIsAuto Property (Word)
keywords: vbawd10.chm26345659
f1_keywords:
- vbawd10.chm26345659
ms.prod: word
api_name:
- Word.Trendline.InterceptIsAuto
ms.assetid: 71abda4e-9de5-71a0-1f0c-f7f81d7e024c
ms.date: 06/08/2017
---


# Trendline.InterceptIsAuto Property (Word)

 **True** if the point where the trendline crosses the value axis is automatically determined by the regression. Read/write **Boolean** .


## Syntax

 _expression_ . **InterceptIsAuto**

 _expression_ A variable that represents a **[Trendline](trendline-object-word.md)** object.


## Remarks

Setting the  **[Intercept](trendline-intercept-property-word.md)** property sets this property to **False** .


## Example

The following example sets Microsoft Word to automatically determine the trendline intercept point for the first chart in the active document. You should run the example on a 2-D column chart that contains a single series that has a trendline.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).Trendlines(1) _ 
 .InterceptIsAuto = True 
 End If 
End With
```


## See also


#### Concepts


[Trendline Object](trendline-object-word.md)

