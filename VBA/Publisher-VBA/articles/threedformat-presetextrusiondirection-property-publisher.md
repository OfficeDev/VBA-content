---
title: "Свойство ThreeDFormat.PresetExtrusionDirection (издатель)"
keywords: vbapb10.chm3801348
f1_keywords: vbapb10.chm3801348
ms.prod: publisher
api_name: Publisher.ThreeDFormat.PresetExtrusionDirection
ms.assetid: fdf3843e-12bc-4b3b-11cb-e512abd991af
ms.date: 06/08/2017
ms.openlocfilehash: d34377eebc13382e839ff7ba4c0b67fe808478c7
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="threedformatpresetextrusiondirection-property-publisher"></a>Свойство ThreeDFormat.PresetExtrusionDirection (издатель)

Возвращает константу **MsoPresetExtrusionDirection** , представляющий путь очистки придания объема от вытянутый фигуры (лицевой из изменяется), занимаемых направление. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PresetExtrusionDirection**

 переменная _expression_A, представляет собой объект- **ThreeDFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoPresetExtrusionDirection


## <a name="remarks"></a>Заметки

Значение свойства **PresetExtrusionDirection** может иметь одно из **MsoPresetExtrusionDirection** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



| **msoExtrusionBottom**|| **msoExtrusionBottomLeft**|| **msoExtrusionBottomRight**|| **msoExtrusionLeft**|| **msoExtrusionNone**|| **msoExtrusionRight**|| **msoExtrusionTop**|| **msoExtrusionTopLeft**|| **msoExtrusionTopRight**|| **msoPresetExtrusionDirectionMixed**| Это свойство доступно только для чтения. Чтобы задать значение этого свойства, используйте метод **[SetExtrusionDirection](threedformat-setextrusiondirection-method-publisher.md)** .


## <a name="example"></a>Пример

В этом примере изменяется изменяется для первой фигуры на первой странице active публикации, если изменяется расширяет направить в верхний левый угол придания объема лицевой. В данном примере для работы указанного фигуры должен быть объемной фигуры.


```vb
Sub SetExtrusion() 
 With ActiveDocument.Pages(1).Shapes(1).ThreeD 
 If .PresetExtrusionDirection = msoExtrusionTopLeft Then 
 .SetExtrusionDirection msoExtrusionBottomRight 
 End If 
 End With 
End Sub
```


