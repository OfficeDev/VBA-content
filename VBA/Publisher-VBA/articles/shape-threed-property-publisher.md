---
title: "Свойство Shape.ThreeD (издатель)"
keywords: vbapb10.chm2228305
f1_keywords: vbapb10.chm2228305
ms.prod: publisher
api_name: Publisher.Shape.ThreeD
ms.assetid: e3430bb2-2f2a-14a6-8eb4-98a29a96ad1c
ms.date: 06/08/2017
ms.openlocfilehash: c5fabec9caabb2091be6f202b8865a5642e41775
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapethreed-property-publisher"></a>Свойство Shape.ThreeD (издатель)

Возвращает объект **[ThreeDFormat](threedformat-object-publisher.md)** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **ThreeD**

 переменная _expression_A, представляющий объект **фигуры** .


## <a name="remarks"></a>Заметки

Свойство **ThreeD** возвращает объект **ThreeDFormat** , свойства которого используются для форматирования объемных внешний вид указанного фигуры.


## <a name="example"></a>Пример

В этом примере задается число уровней, цвет объемной фигуры, направление придания объема и направление освещения для объемных эффектов, примененных к фигуры один активный публикации.


```vb
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

```


