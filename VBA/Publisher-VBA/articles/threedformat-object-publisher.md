---
title: "Объект ThreeDFormat (издатель)"
keywords: vbapb10.chm3866623
f1_keywords: vbapb10.chm3866623
ms.prod: publisher
api_name: Publisher.ThreeDFormat
ms.assetid: 11d57330-c99e-5aa9-d47c-2c5d2846ed4d
ms.date: 06/08/2017
ms.openlocfilehash: afbab5d0a217f805150510ac0a7b7da62049a567
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="threedformat-object-publisher"></a>Объект ThreeDFormat (издатель)

Представляет трехмерную форматирования фигуры.
 


## <a name="remarks"></a>Заметки

Не удается применить трехмерную форматирование некоторые виды фигур, таких как среза фигур. Большая часть свойств и методов объекта **ThreeDFormat** для таких фигуры завершится с ошибкой.
 

 

## <a name="example"></a>Пример

Свойство **[ThreeD](shape-threed-property-publisher.md)** используется для возврата объекта **ThreeDFormat** . В этом примере задается число уровней, цвет объемной фигуры, направление придания объема и направление освещения для объемных эффектов, примененных к фигуры один активный публикации.
 

 

```
Sub SetThreeDSettings() 
 Dim tdfTemp As ThreeDFormat 
 
 Set tdfTemp = _ 
 ActiveDocument.Pages(1).Shapes(1).ThreeD 
 
 With tdfTemp 
 .Visible = True 
 .Depth = 50 
 .ExtrusionColor.RGB = RGB(255, 100, 255) 
 .SetExtrusionDirection _ 
 PresetExtrusionDirection:=msoExtrusionTop 
 .PresetLightingDirection = msoLightingLeft 
 End With 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[IncrementRotationX](threedformat-incrementrotationx-method-publisher.md)|
|[IncrementRotationY](threedformat-incrementrotationy-method-publisher.md)|
|[ResetRotation](threedformat-resetrotation-method-publisher.md)|
|[SetExtrusionDirection](threedformat-setextrusiondirection-method-publisher.md)|
|[SetThreeDFormat](threedformat-setthreedformat-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](threedformat-application-property-publisher.md)|
|[BevelBottomDepth](threedformat-bevelbottomdepth-property-publisher.md)|
|[BevelBottomInset](threedformat-bevelbottominset-property-publisher.md)|
|[BevelBottomType](threedformat-bevelbottomtype-property-publisher.md)|
|[BevelTopDepth](threedformat-beveltopdepth-property-publisher.md)|
|[BevelTopInset](threedformat-beveltopinset-property-publisher.md)|
|[BevelTopType](threedformat-beveltoptype-property-publisher.md)|
|[ContourColor](threedformat-contourcolor-property-publisher.md)|
|[ContourWidth](threedformat-contourwidth-property-publisher.md)|
|[Число уровней](threedformat-depth-property-publisher.md)|
|[ExtrusionColor](threedformat-extrusioncolor-property-publisher.md)|
|[ExtrusionColorType](threedformat-extrusioncolortype-property-publisher.md)|
|[FieldOfView](threedformat-fieldofview-property-publisher.md)|
|[Родительский раздел](threedformat-parent-property-publisher.md)|
|[Перспектива](threedformat-perspective-property-publisher.md)|
|[PresetExtrusionDirection](threedformat-presetextrusiondirection-property-publisher.md)|
|[PresetLightingDirection](threedformat-presetlightingdirection-property-publisher.md)|
|[PresetLightingSoftness](threedformat-presetlightingsoftness-property-publisher.md)|
|[PresetMaterial](threedformat-presetmaterial-property-publisher.md)|
|[PresetThreeDFormat](threedformat-presetthreedformat-property-publisher.md)|
|[RotationX](threedformat-rotationx-property-publisher.md)|
|[RotationY](threedformat-rotationy-property-publisher.md)|
|[Visible](threedformat-visible-property-publisher.md)|

