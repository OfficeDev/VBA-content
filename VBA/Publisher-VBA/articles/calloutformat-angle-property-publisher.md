---
title: "Свойство CalloutFormat.Angle (издатель)"
keywords: vbapb10.chm2490625
f1_keywords: vbapb10.chm2490625
ms.prod: publisher
api_name: Publisher.CalloutFormat.Angle
ms.assetid: b65a1c87-db52-8703-135e-1fbb1efbeebe
ms.date: 06/08/2017
ms.openlocfilehash: fa002f8fb9b55b60f0bd4db326e4de881f204c63
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="calloutformatangle-property-publisher"></a>Свойство CalloutFormat.Angle (издатель)

Возвращает или задает константой **MsoCalloutAngleType** , представляющее угол линии выноски. Если строка выноски содержит более одного сегмента линии, это свойство Возвращает или задает угол сегмент, дальше всего от текстовое поле выноски. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Угол**

 переменная _expression_A, представляет собой объект- **CalloutFormat** .


## <a name="remarks"></a>Заметки

Если задать значение этого свойства отличное от **msoCalloutAngleAutomatic**, линии выноски поддерживает фиксированный угол при перетаскивании выноске.



| MsoCalloutAngleType может иметь одно из следующих констант MsoCalloutAngleType. | | **msoCalloutAngle30**|| **msoCalloutAngle45**|| **msoCalloutAngle60**|| **msoCalloutAngle90**|| **msoCalloutAngleAutomatic**|| **msoCalloutAngleMixed**|

## <a name="example"></a>Пример

В этом примере задается угол выноски до 90 градусов для первой фигуры на первой странице active публикации. В данном примере для работы указанного фигуры должен быть выноске.


```vb
Sub SetCalloutAngle() 
 ActiveDocument.Pages(1).Shapes(1).Callout.Angle = msoCalloutAngle90 
End Sub
```


