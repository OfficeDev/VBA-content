---
title: "Свойство ShadowFormat.OffsetY (издатель)"
keywords: vbapb10.chm3670275
f1_keywords: vbapb10.chm3670275
ms.prod: publisher
api_name: Publisher.PictureFormat.OffsetY
ms.assetid: e7deb108-e027-dd61-714f-1a76e904009b
ms.date: 06/08/2017
ms.openlocfilehash: 31d6466a1f24622e8cbedb193153a79ab53bdb90
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shadowformatoffsety-property-publisher"></a>Свойство ShadowFormat.OffsetY (издатель)

Возвращает или задает значение **Variant** , указывающее горизонтальное смещение тени заданной фигуры. Положительное значение смещения тени справа от фигуры; отрицательное значение смещения его слева. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **OffsetY**

 переменная _expression_A, представляющий объект **ShadowFormat** .


### <a name="return-value"></a>Возвращаемое значение

Variant


## <a name="remarks"></a>Заметки

Числовые значения вычисляются в точках; строк может быть в любой устройств, поддерживаемых Microsoft Publisher (например, «2,5 дюйма»).

Если вы хотите Сдвиг тени по горизонтали или по вертикали из текущей позиции без указания абсолютного положения, используйте метод **[IncrementOffsetX](shadowformat-incrementoffsetx-method-publisher.md)** или **[IncrementOffsetY](shadowformat-incrementoffsety-method-publisher.md)** .


## <a name="example"></a>Пример

В этом примере задается горизонтального и вертикального смещения тени для трех фигуры на странице один активный публикации. 5 точек справа от фигуры и 3 точки над текстом смещения тени. Если фигуры еще нет тени, этот пример добавляет в него.


```vb
With ActiveDocument.Pages(1).Shapes(3).Shadow 
 .Visible = True 
 .OffsetX = 5 
 .OffsetY = -3 
End With
```


