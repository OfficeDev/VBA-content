---
title: "Свойство Shape.ZOrderPosition (издатель)"
keywords: vbapb10.chm2228312
f1_keywords: vbapb10.chm2228312
ms.prod: publisher
api_name: Publisher.Shape.ZOrderPosition
ms.assetid: 46eb765b-578e-f6df-43b7-c14443cddbb2
ms.date: 06/08/2017
ms.openlocfilehash: 0bd9cde9245460e5c9d7b0276c64e592f2440141
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapezorderposition-property-publisher"></a>Свойство Shape.ZOrderPosition (издатель)

Возвращает значение типа **Long** , указывающее положение указанные форму или диапазона фигуры в z порядке. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ZOrderPosition**

 переменная _expression_A, представляющий объект **фигуры** .


## <a name="remarks"></a>Заметки

Номер индекса фигуры в коллекции **фигур** соответствует фигуры позицию в z порядке. Например, если существует четыре фигуры на странице выражение `ActiveDocument.Pages(1).Shapes(1)` возвращает фигуры на задней z порядка и выражение `ActiveDocument.Pages(1).Shapes(4)` возвращает форму в начале z порядке.

При добавлении новой фигуры в семейство сайтов по умолчанию будет добавлена в начало z порядка.

Чтобы задать положение фигуры в z порядке, используйте метод [ZOrder](shape-zorder-method-publisher.md) .


## <a name="example"></a>Пример

В этом примере добавляет овала active публикации и помещает Овал второй с обратной в z порядке при наличии по крайней мере один фигуры на странице.


```vb
With ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShapeOval, _ 
 Left:=100, Top:=100, Width:=100, Height:=300) 
 Do While .ZOrderPosition > 2 
 .ZOrder msoSendBackward 
 Loop 
End With 

```


