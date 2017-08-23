---
title: "Объект FillFormat (издатель)"
keywords: vbapb10.chm2424831
f1_keywords: vbapb10.chm2424831
ms.prod: publisher
api_name: Publisher.FillFormat
ms.assetid: 0a5d4f7a-c42a-28ad-c86d-ac9828a3b874
ms.date: 06/08/2017
ms.openlocfilehash: 2240a9c5d4f6367a4724db065d983dfdbc07b372
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fillformat-object-publisher"></a>Объект FillFormat (издатель)

Представляет заполнения форматирования для фигуры. Фигура может иметь сплошной, градиентной, текстуры, шаблон, рисунков или Полупрозрачная заливки.
 


## <a name="remarks"></a>Заметки

Многие из свойств объекта **FillFormat** доступны только для чтения. Чтобы задать одно из этих свойств, необходимо применить соответствующий метод.
 

 

## <a name="example"></a>Пример

Используйте свойство **[заполните поля](shape-fill-property-publisher.md)** для возврата объекта **FillFormat** . В следующем примере добавляется фигура в активный документ и затем задает градиентной и цвет заливки фигуры.
 

 

```
Sub AddShapeAndSetFill() 
 With ActiveDocument.Pages(1).Shapes.AddShape(Type:=msoShapeHeart, _ 
 Left:=90, Top:=90, Width:=90, Height:=80).Fill 
 .ForeColor.RGB = RGB(Red:=255, Green:=0, Blue:=0) 
 .OneColorGradient Style:=msoGradientHorizontal, _ 
 Variant:=1, Degree:=1 
 End With 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[OneColorGradient](fillformat-onecolorgradient-method-publisher.md)|
|[Узорные](fillformat-patterned-method-publisher.md)|
|[PresetGradient](fillformat-presetgradient-method-publisher.md)|
|[PresetTextured](fillformat-presettextured-method-publisher.md)|
|[Сплошной](fillformat-solid-method-publisher.md)|
|[TwoColorGradient](fillformat-twocolorgradient-method-publisher.md)|
|[UserPicture](fillformat-userpicture-method-publisher.md)|
|[UserTextured](fillformat-usertextured-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](fillformat-application-property-publisher.md)|
|[Цвет фона](fillformat-backcolor-property-publisher.md)|
|[Цвет текста](fillformat-forecolor-property-publisher.md)|
|[GradientAngle](fillformat-gradientangle-property-publisher.md)|
|[GradientColorType](fillformat-gradientcolortype-property-publisher.md)|
|[GradientDegree](fillformat-gradientdegree-property-publisher.md)|
|[GradientStyle](fillformat-gradientstyle-property-publisher.md)|
|[GradientVariant](fillformat-gradientvariant-property-publisher.md)|
|[Родительский раздел](fillformat-parent-property-publisher.md)|
|[Шаблон](fillformat-pattern-property-publisher.md)|
|[PresetGradientType](fillformat-presetgradienttype-property-publisher.md)|
|[PresetTexture](fillformat-presettexture-property-publisher.md)|
|[RotateWithObject](fillformat-rotatewithobject-property-publisher.md)|
|[TextureAlignment](fillformat-texturealignment-property-publisher.md)|
|[TextureHorizontalScale](fillformat-texturehorizontalscale-property-publisher.md)|
|[TextureName](fillformat-texturename-property-publisher.md)|
|[TextureOffsetX](fillformat-textureoffsetx-property-publisher.md)|
|[TextureOffsetY](fillformat-textureoffsety-property-publisher.md)|
|[TextureType](fillformat-texturetype-property-publisher.md)|
|[TextureVerticalScale](fillformat-textureverticalscale-property-publisher.md)|
|[Прозрачность](fillformat-transparency-property-publisher.md)|
|[Type](fillformat-type-property-publisher.md)|
|[Visible](fillformat-visible-property-publisher.md)|

