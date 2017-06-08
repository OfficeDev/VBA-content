---
title: FillFormat Object (PowerPoint)
keywords: vbapp10.chm552000
f1_keywords:
- vbapp10.chm552000
ms.prod: powerpoint
api_name:
- PowerPoint.FillFormat
ms.assetid: 5bd4e2cb-4466-b468-d494-bec30ed5c9d8
ms.date: 06/08/2017
---


# FillFormat Object (PowerPoint)

Represents fill formatting for a shape. A shape can have a solid, gradient, texture, pattern, picture, or semi-transparent fill.


## Remarks

Many of the properties of the  **FillFormat** object are read-only. To set one of these properties, you have to apply the corresponding method.


## Example

Use the  **Fill** property to return a **FillFormat** object. The following example adds a rectangle to `myDocument` and then sets the gradient and color for the rectangle's fill.


```
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes _

        .AddShape(msoShapeRectangle, 90, 90, 90, 80).Fill

    .ForeColor.RGB = RGB(0, 128, 128)

    .OneColorGradient msoGradientHorizontal, 1, 1

End With
```


## Methods



|**Name**|
|:-----|
|[Background](http://msdn.microsoft.com/library/4c82e3d3-86cd-d18f-ead1-9fc2dda5efd8%28Office.15%29.aspx)|
|[OneColorGradient](http://msdn.microsoft.com/library/ce574185-2d13-993b-4a78-d681b6600621%28Office.15%29.aspx)|
|[Patterned](http://msdn.microsoft.com/library/665c5b1d-e2a2-64ab-a0c3-7d22d8d3121a%28Office.15%29.aspx)|
|[PresetGradient](http://msdn.microsoft.com/library/6aa304c7-a2ee-ceea-f956-404538bebc43%28Office.15%29.aspx)|
|[PresetTextured](http://msdn.microsoft.com/library/a025a1d3-a2db-e219-7080-1a29c2fd3f21%28Office.15%29.aspx)|
|[Solid](http://msdn.microsoft.com/library/0d3302de-2b8b-2a05-697d-0010882588e5%28Office.15%29.aspx)|
|[TwoColorGradient](http://msdn.microsoft.com/library/29dac3d9-366e-0fd5-0fe3-dc64fa2fc871%28Office.15%29.aspx)|
|[UserPicture](http://msdn.microsoft.com/library/87f28942-a5d2-7e27-7eee-5181d112d6d2%28Office.15%29.aspx)|
|[UserTextured](http://msdn.microsoft.com/library/351d00db-4ed3-6975-e9c6-4174e796395d%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/5ef3dc11-eddf-48a7-4cf6-64149b0bf903%28Office.15%29.aspx)|
|[BackColor](http://msdn.microsoft.com/library/d78fa88b-578d-f469-f2e1-7564ebc91f8d%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/f2d09239-4438-ac63-41d6-414cda762802%28Office.15%29.aspx)|
|[ForeColor](http://msdn.microsoft.com/library/3dc07a0f-d0bc-52c8-e06a-dd0315151742%28Office.15%29.aspx)|
|[GradientAngle](http://msdn.microsoft.com/library/eb5362f0-5d3b-0091-7a83-0a8d58d90438%28Office.15%29.aspx)|
|[GradientColorType](http://msdn.microsoft.com/library/90224ee2-80f9-480b-bd1b-678035ded3ef%28Office.15%29.aspx)|
|[GradientDegree](http://msdn.microsoft.com/library/201380df-f7b4-a38c-e615-2eb490b7042c%28Office.15%29.aspx)|
|[GradientStops](http://msdn.microsoft.com/library/dd0c2c5a-81f1-b008-5b2f-5248241ac0db%28Office.15%29.aspx)|
|[GradientStyle](http://msdn.microsoft.com/library/dca37bf2-1219-d815-7584-97a8665e3420%28Office.15%29.aspx)|
|[GradientVariant](http://msdn.microsoft.com/library/32a8a1fd-84aa-fbee-35c5-5bd83b0790c6%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/b81440f3-aa91-7a67-0a61-a30cf40e2c29%28Office.15%29.aspx)|
|[Pattern](http://msdn.microsoft.com/library/843504d6-d9a5-f732-89eb-d2d3d1ea4477%28Office.15%29.aspx)|
|[PictureEffects](http://msdn.microsoft.com/library/01897ad5-84c9-f98e-8c2f-9a9e5c13bc2e%28Office.15%29.aspx)|
|[PresetGradientType](http://msdn.microsoft.com/library/a9a4f3fc-7350-aba1-394a-10936166ea4c%28Office.15%29.aspx)|
|[PresetTexture](http://msdn.microsoft.com/library/684d39f9-53d8-4f69-a6ae-c447253ae3a7%28Office.15%29.aspx)|
|[RotateWithObject](http://msdn.microsoft.com/library/46197f92-b12a-957f-1ab5-063b0d4d2933%28Office.15%29.aspx)|
|[TextureAlignment](http://msdn.microsoft.com/library/e26ca83c-7dc1-4c7b-52a4-3a30669079ea%28Office.15%29.aspx)|
|[TextureHorizontalScale](http://msdn.microsoft.com/library/3ffaf1b9-0657-96b4-9c28-39c111200f1d%28Office.15%29.aspx)|
|[TextureName](http://msdn.microsoft.com/library/c8ca47e7-90c8-50b8-2e7e-29e56ec0f70e%28Office.15%29.aspx)|
|[TextureOffsetX](http://msdn.microsoft.com/library/5c0a5dd6-ff18-6094-7e27-0dfe934f2028%28Office.15%29.aspx)|
|[TextureOffsetY](http://msdn.microsoft.com/library/f1ba83a3-65ca-dd4c-cb70-f6cb453b824c%28Office.15%29.aspx)|
|[TextureTile](http://msdn.microsoft.com/library/14d1b329-8d06-b4d6-1ade-aea80f5427ce%28Office.15%29.aspx)|
|[TextureType](http://msdn.microsoft.com/library/318e5b2f-7baa-296b-c7ea-0feddb70414c%28Office.15%29.aspx)|
|[TextureVerticalScale](http://msdn.microsoft.com/library/714f17bd-db5b-4b09-c166-69f25e7a59d5%28Office.15%29.aspx)|
|[Transparency](http://msdn.microsoft.com/library/98b099d7-9149-d306-1a80-f85b89b029c5%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/e7818487-0e6f-3227-487d-94ffeaf85006%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/8221347f-4b12-f18a-5d0b-b584ee762bff%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
