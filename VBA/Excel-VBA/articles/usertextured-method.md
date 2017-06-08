---
title: UserTextured Method
keywords: vbagr10.chm5208113
f1_keywords:
- vbagr10.chm5208113
ms.prod: excel
api_name:
- Excel.UserTextured
ms.assetid: 063b74ef-8b82-3a59-457c-9240395a6eb2
ms.date: 06/08/2017
---


# UserTextured Method

Fills the specified shape with small tiles of an image. If you want to fill the shape with one large image, use the  **UserPicture** method.

 _expression_. **UserTextured**( **_TextureFile_**)

 _expression_ Required. An expression that returns a **ChartFillFormat** object.

 **TextureFile** Required **String**. The name of the specified picture file.

## Example

This example changes the user-defined texture type for the chart's fill format.


```vb
With myChart.ChartArea.Fill 
 If .Type = msoFillTextured Then 
 If .TextureType = msoTextureUserDefined Then 
 If .TextureName = "C:\brick.bmp" Then 
 .UserTextured "C:\stone.bmp" 
 End If 
 End If 
 End If 
End With
```


