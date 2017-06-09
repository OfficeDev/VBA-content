---
title: TextureType Property
keywords: vbagr10.chm3077593
f1_keywords:
- vbagr10.chm3077593
ms.prod: excel
api_name:
- Excel.TextureType
ms.assetid: ba60a953-c506-ff49-0945-aa222dcd5f43
ms.date: 06/08/2017
---


# TextureType Property

Returns the texture type for the specified fill. Read-only MsoTextureType .



|MsoTextureType can be one of these MsoTextureType constants.|
| **msoTexturePreset**|
| **msoTextureTypeMixed**|
| **msoTextureUserDefined**This property is read-only. Use the  **UserTextured** method to set the texture type for the fill.|

 _expression_. **TextureType**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

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


