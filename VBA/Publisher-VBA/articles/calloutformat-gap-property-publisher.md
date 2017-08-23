---
title: "Свойство CalloutFormat.Gap (издатель)"
keywords: vbapb10.chm2490631
f1_keywords: vbapb10.chm2490631
ms.prod: publisher
api_name: Publisher.CalloutFormat.Gap
ms.assetid: fd7cdac7-5f09-a574-e9ef-08feebd81cff
ms.date: 06/08/2017
ms.openlocfilehash: b93da27408e81f94ea98f10d2645f303147096de
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="calloutformatgap-property-publisher"></a>Свойство CalloutFormat.Gap (издатель)

Возвращает или задает **Variant** , указывающее расстояние по горизонтали между в конец строки выноски и текст, ограничивающий прямоугольник. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Разрывов**

 переменная _expression_A, представляет собой объект- **CalloutFormat** .


### <a name="return-value"></a>Возвращаемое значение

Variant


## <a name="remarks"></a>Заметки

Числовые значения вычисляются в точках; строк может быть в любой устройств, поддерживаемых Microsoft Publisher (например, «2,5 дюйма»).


## <a name="example"></a>Пример

В этом примере задается расстояние между линии выноски и ограничивающий текстовое поле 3 пунктов для первой фигуры в активной публикации. Для обеспечения работы примера фигуры должен быть выноске.


```vb
ActiveDocument.Pages(1).Shapes(1).Callout.Gap = 3
```


