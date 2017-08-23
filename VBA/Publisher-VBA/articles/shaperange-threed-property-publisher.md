---
title: "Свойство ShapeRange.ThreeD (издатель)"
keywords: vbapb10.chm2293841
f1_keywords: vbapb10.chm2293841
ms.prod: publisher
api_name: Publisher.ShapeRange.ThreeD
ms.assetid: e5905f9d-dd84-b97e-ac5d-630f6c1208d7
ms.date: 06/08/2017
ms.openlocfilehash: c5ac3603fc39abde748a9d9afffc0772b25dba40
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangethreed-property-publisher"></a>Свойство ShapeRange.ThreeD (издатель)

Возвращает объект **[ThreeDFormat](threedformat-object-publisher.md)** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **ThreeD**

 переменная _expression_A, представляющий объект **ShapeRange** .


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


