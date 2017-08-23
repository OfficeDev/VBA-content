---
title: "Метод ThreeDFormat.IncrementRotationY (издатель)"
keywords: vbapb10.chm3801105
f1_keywords: vbapb10.chm3801105
ms.prod: publisher
api_name: Publisher.ThreeDFormat.IncrementRotationY
ms.assetid: 54260253-c914-6600-60ef-17bdde12be59
ms.date: 06/08/2017
ms.openlocfilehash: 7a608580354ead9fb2ab2201d3effc63cee0b15c
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="threedformatincrementrotationy-method-publisher"></a>Метод ThreeDFormat.IncrementRotationY (издатель)

Изменяет вращение указанного фигуры относительно оси y (по вертикали) указанное число градусов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **IncrementRotationY** ( **_Порядкового номера_**)

 переменная _expression_A, представляет собой объект- **ThreeDFormat** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Порядкового номера|Обязательное свойство.| **Один**|Указывает, сколько градусов Поворот фигуры относительно оси y. Может быть в диапазоне от - 90 до 90. Положительное значение tilts фигуры слева; отрицательное значение tilts его справа.|

## <a name="remarks"></a>Заметки

Свойство **[RotationY](threedformat-rotationy-property-publisher.md)** задать абсолютные вращения фигуры относительно оси y.

Нельзя изменять размеры поворот вокруг оси y указанной фигуры за границу верхней или нижней, для свойства **RotationY** (90 градусов - 90 градусов). К примеру Если свойство **RotationY** изначально установлено значение 80 и можно указать 40 для аргумента **_порядкового номера_** , результирующий вращение 90 (верхний предел для свойства **RotationY** ) вместо 120.

Чтобы изменить вращения фигуры относительно оси x (по горизонтали), используйте метод **[IncrementRotationX](threedformat-incrementrotationx-method-publisher.md)** . Чтобы изменить вращения вокруг оси z (расширяет наружу плоскости публикации), используйте метод **[IncrementRotation](shape-incrementrotation-method-publisher.md)** .


## <a name="example"></a>Пример

В этом примере tilts первую фигуру в активной публикации 10 градусов вправо. Фигуры должен быть вытянутый фигуру увидеть результат этого кода.


```vb
ActiveDocument.Pages(1).Shapes(1).ThreeD _ 
 .IncrementRotationY Increment:=-10
```


