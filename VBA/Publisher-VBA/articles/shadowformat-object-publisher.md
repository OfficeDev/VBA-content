---
title: "Объект ShadowFormat (издатель)"
keywords: vbapb10.chm3735551
f1_keywords: vbapb10.chm3735551
ms.prod: publisher
api_name: Publisher.ShadowFormat
ms.assetid: b23ab92e-5e49-8d8d-69d5-93d391a9edb2
ms.date: 06/08/2017
ms.openlocfilehash: 484ae165ecbc42f5c6c367ed68818200ecab4e42
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shadowformat-object-publisher"></a>Объект ShadowFormat (издатель)

Представляет тени для фигуры.
 


## <a name="example"></a>Пример

Используйте свойство **теневой** возвращает объект **ShadowFormat** . Следующий пример добавляет тенью прямоугольника в активный документ. Смещения розовым тени 7 точки справа от прямоугольника и 7 точки над текстом.
 

 

```
Sub FormatShadow() 
 With ActiveDocument.Pages(1).Shapes.AddShape( _ 
 Type:=msoShapeRectangle, Left:=72, Top:=72, _ 
 Width:=100, Height:=200).Shadow 
 .ForeColor.RGB = RGB(Red:=255, Green:=0, Blue:=150) 
 .Obscured = msoTrue 
 .OffsetX = 7 
 .OffsetY = -7 
 .Visible = True 
 End With 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[IncrementOffsetX](shadowformat-incrementoffsetx-method-publisher.md)|
|[IncrementOffsetY](shadowformat-incrementoffsety-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](shadowformat-application-property-publisher.md)|
|[Чтобы минимизировать](shadowformat-blur-property-publisher.md)|
|[Цвет текста](shadowformat-forecolor-property-publisher.md)|
|[Закрыты](shadowformat-obscured-property-publisher.md)|
|[OffsetX](shadowformat-offsetx-property-publisher.md)|
|[OffsetY](shadowformat-offsety-property-publisher.md)|
|[Родительский раздел](shadowformat-parent-property-publisher.md)|
|[RotateWithShape](shadowformat-rotatewithshape-property-publisher.md)|
|[Размер](shadowformat-size-property-publisher.md)|
|[Type](shadowformat-type-property-publisher.md)|
|[Visible](shadowformat-visible-property-publisher.md)|

