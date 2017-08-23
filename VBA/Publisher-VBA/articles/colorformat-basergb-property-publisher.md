---
title: "Свойство ColorFormat.BaseRGB (издатель)"
keywords: vbapb10.chm2555906
f1_keywords: vbapb10.chm2555906
ms.prod: publisher
api_name: Publisher.ColorFormat.BaseRGB
ms.assetid: c8096661-9a5a-2769-fd88-72d38d383095
ms.date: 06/08/2017
ms.openlocfilehash: 3614acab7b04ac51a2d54988cf84a8c3d2a6e5d9
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="colorformatbasergb-property-publisher"></a>Свойство ColorFormat.BaseRGB (издатель)

Возвращает или задает константой **MsoRGBType** , который представляет исходный формат цвета RGB перед изменением цвета свойства, такие как оттенок и тени, которые применяются. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **BaseRGB**

 переменная _expression_A, представляет собой объект- **ColorFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoRGBType


## <a name="example"></a>Пример

В этом примере создается фигуры, задает цвет заливки и осветляет цвета; Затем он создает второй фигуры и применение исходного цвета RGB первой фигуры для второй фигуры.


```vb
Sub SetBaseRGB() 
 Dim shpOne As Shape 
 
 With ActiveDocument.Pages(1).Shapes 
 Set shpOne = .AddShape(Type:=msoShapeHeart, _ 
 Left:=150, Top:=150, Width:=300, Height:=300) 
 With shpOne.Fill.ForeColor 
 .RGB = RGB(Red:=160, Green:=0, Blue:=255) 
 .TintAndShade = 0.9 
 End With 
 .AddShape(Type:=msoShapeRectangle, Left:=62, _ 
 Top:=500, Width:=488, Height:=100).Fill _ 
 .ForeColor.RGB = shpOne.Fill.ForeColor.BaseRGB 
 End With 
End Sub
```


