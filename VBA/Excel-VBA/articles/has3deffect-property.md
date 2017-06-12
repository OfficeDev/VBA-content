---
title: Has3DEffect Property
keywords: vbagr10.chm67201
f1_keywords:
- vbagr10.chm67201
ms.prod: excel
api_name:
- Excel.Has3DEffect
ms.assetid: e19f4d47-ca7b-ea70-01eb-ced3c1dd343f
ms.date: 06/08/2017
---


# Has3DEffect Property

 **True** if the series has a three-dimensional appearance. Applies only to bubble charts. Read/write **Boolean**.


## Example

This example gives series one on the bubble chart a three-dimensional appearance.


```vb
With myChart 
 .SeriesCollection(1).Has3DEffect = True 
End With
```


