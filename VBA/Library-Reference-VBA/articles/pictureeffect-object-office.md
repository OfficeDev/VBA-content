---
title: PictureEffect Object (Office)
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.PictureEffect
ms.assetid: af3f742a-e082-1abd-7df2-d1fb2f57c8a2
---


# PictureEffect Object (Office)

Represents a Picture Effect.


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
|[Delete](http://msdn.microsoft.com/library/cd107111-0866-fa75-bdbf-6a0cc562c815%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/90e612f1-71b6-48d7-4c14-0336d0992cc3%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/4d001927-b503-34a9-0776-bb186a22cb96%28Office.15%29.aspx)|
|[EffectParameters](http://msdn.microsoft.com/library/a0729015-14ab-e5c3-9772-678b892e4834%28Office.15%29.aspx)|
|[Position](http://msdn.microsoft.com/library/29c2d136-777f-5984-3018-3dae2721ed76%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/9d93d9b5-726b-5cbb-3642-bbd461d706c7%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/cdfcda14-5d74-c61f-e289-1d53ea3e8e80%28Office.15%29.aspx)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
[PictureEffect Object Members](http://msdn.microsoft.com/library/df7a24cd-db6f-1ab1-e0e4-3b332ba27bd5%28Office.15%29.aspx)
