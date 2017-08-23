---
title: "Свойство ThreeDFormat.PresetLightingDirection (издатель)"
keywords: vbapb10.chm3801349
f1_keywords: vbapb10.chm3801349
ms.prod: publisher
api_name: Publisher.ThreeDFormat.PresetLightingDirection
ms.assetid: 94957653-a4e1-bcb6-7697-ed10d1b54301
ms.date: 06/08/2017
ms.openlocfilehash: 7cd170da7ebda6245bd99236ba98ad44babe64a8
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="threedformatpresetlightingdirection-property-publisher"></a>Свойство ThreeDFormat.PresetLightingDirection (издатель)

Возвращает или задает константой **MsoPresetLightingDirection** , представляющий позиции источника света относительно изменяется. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PresetLightingDirection**

 переменная _expression_A, представляет собой объект- **ThreeDFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoPresetLightingDirection


## <a name="remarks"></a>Заметки

Значение свойства **PresetLightingDirection** может иметь одно из ** [MsoPresetLightingDirection](http://msdn.microsoft.com/library/d3de37f8-f4c8-d04f-12a9-5fb7340fb8b1%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.

Не будет видно, если изменяется поверхность каркас световые эффекты, заданную вами.


## <a name="example"></a>Пример

В этом примере задается изменяется для первой фигуры на первой странице active публикации для расширения к началу фигуры и освещения для изменяется поступают из слева. В данном примере для работы указанного фигуры должен быть объемной фигуры.


```vb
Sub ExtrusionLighting() 
 With ActiveDocument.Pages(1).Shapes(1).ThreeD 
 .Visible = True 
 .SetExtrusionDirection msoExtrusionTop 
 .PresetLightingDirection = msoLightingLeft 
 End With 
End Sub
```


