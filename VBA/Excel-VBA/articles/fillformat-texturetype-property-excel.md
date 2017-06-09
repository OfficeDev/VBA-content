---
title: FillFormat.TextureType Property (Excel)
keywords: vbaxl10.chm115021
f1_keywords:
- vbaxl10.chm115021
ms.prod: excel
api_name:
- Excel.FillFormat.TextureType
ms.assetid: 9a39c34e-c19c-5539-b5ac-b624fe71e2e9
ms.date: 06/08/2017
---


# FillFormat.TextureType Property (Excel)

Returns the texture type for the specified fill. Read-only  **[MsoTextureType](http://msdn.microsoft.com/library/be7fdbb6-3684-fa23-f1d8-f0caac02754e%28Office.15%29.aspx)** .


## Syntax

 _expression_ . **TextureType**

 _expression_ A variable that represents a **FillFormat** object.


## Remarks

Use the  **[UserTextured](fillformat-usertextured-method-excel.md)** method to set the texture type for the fill.


## Example

This example sets the fill format for chart two to the same style used for chart one.


```vb
Set c1f = Charts(1).ChartArea.Fill 
If c1f.Type = msoFillTextured Then 
    With Charts(2).ChartArea.Fill 
        .Visible = True 
        If c1f.TextureType = msoTexturePreset Then 
            .PresetTextured c1f.PresetTexture 
        Else 
            .UserTextured c1f.TextureName 
        End If 
    End With 
End If
```


## See also


#### Concepts


[FillFormat Object](fillformat-object-excel.md)

