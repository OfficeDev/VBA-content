---
title: PictureEffects Object (Office)
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.PictureEffects
ms.assetid: bc0e1cfd-7328-360d-872e-c71ae93162ed
---


# PictureEffects Object (Office)

Represents a collection of  **PictureEffects** objects.


## Remarks

Picture Effects are processed as a chain composed of individual items which are applied in sequence to create the final composited image. An Effects chain will allow an effect to be added to the chain, reordered, or removed from the chain.


## Example

The following code sets several Picture Effect fill properties on a shape in a Microsoft PowerPoint slide.


```
Sub PictureEffectSample() 
' Setup a slide with one picture shape. 
With ActivePresentation.Slides(1).Shapes(1).Fill.PictureEffects 
 
 ' Insert a 150% Saturation effect. 
 .Insert(msoEffectSaturation).EffectParameters(1).Value = 1.5 
 
 ' Insert Brightness/Contrast effect and set values to -50% Brightness and +25% Contrast. 
 Dim brightnessContrast As PictureEffect 
 Set brightnessContrast = .Insert(msoEffectBrightnessContrast) 
 brightnessContrast.EffectParameters(1).Value = -0.5 
 brightnessContrast.EffectParameters(2).Value = 0.25 
 
 ' Remove all Picture effects. 
 While .Count > 0 
 .Delete (1) 
 Wend 
 
End With 
End Sub 

```


## Methods



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/791c409d-26e6-b4d7-8625-ad8cfe7c797e%28Office.15%29.aspx)|
|[Insert](http://msdn.microsoft.com/library/589c38d7-1d0a-ad87-a84c-72147b6b07cf%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/3aa0b57d-2bf7-8d54-3b2e-e23bef0f20af%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/3832dfbd-8c4c-fbee-613d-f31d2b1c9387%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/587a6d8a-9c50-802e-1e10-561c821eb985%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/bd92a68a-059b-d96c-a86f-7c6754b23026%28Office.15%29.aspx)|

## See also


#### Other resources


[PictureEffects Object Members](http://msdn.microsoft.com/library/fe7a9f46-f5fa-8ab9-5fb6-c88d283e4663%28Office.15%29.aspx)
[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
