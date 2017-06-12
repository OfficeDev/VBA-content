---
title: TextureName Property
keywords: vbagr10.chm5208045
f1_keywords:
- vbagr10.chm5208045
ms.prod: excel
api_name:
- Excel.TextureName
ms.assetid: a2c0e2af-5f16-f181-0404-49223de24a97
ms.date: 06/08/2017
---


# TextureName Property

Returns the name of the custom texture file for the specified fill. Read-only  **String**.

This property is read-only. Use the  **UserPicture** or **UserTextured** method to set the texture file for the fill.

## Example

This example changes the user-defined texture type for the chart's fill format.


```vb
With myChart.ChartArea.Fill 
 If .Type = msoFillTextured Then 
 If .TextureType = msoTextureUserDefined Then 
 If .TextureName = "brick.bmp" Then 
 .UserTextured "stone.bmp" 
 End If 
 End If 
 End If 
End With
```


