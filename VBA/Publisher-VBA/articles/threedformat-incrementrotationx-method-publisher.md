---
title: "Метод ThreeDFormat.IncrementRotationX (издатель)"
keywords: vbapb10.chm3801104
f1_keywords: vbapb10.chm3801104
ms.prod: publisher
api_name: Publisher.ThreeDFormat.IncrementRotationX
ms.assetid: d64204d6-ff4e-aa25-7795-858006ba2cf2
ms.date: 06/08/2017
ms.openlocfilehash: f08dff5bfffa2492c457221291a513a49d467706
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="threedformatincrementrotationx-method-publisher"></a>Метод ThreeDFormat.IncrementRotationX (издатель)

Изменяет вращение указанного фигуры вокруг оси x (по горизонтали) указанное число градусов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **IncrementRotationX** ( **_Порядкового номера_**)

 переменная _expression_A, представляет собой объект- **ThreeDFormat** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Порядкового номера|Обязательное свойство.| **Один**|Указывает, сколько градусов Поворот фигуры относительно оси x. Может быть в диапазоне от - 90 до 90. Положительное значение tilts фигуры. отрицательное значение tilts его работу.|

## <a name="remarks"></a>Заметки

Свойство **[RotationX](threedformat-rotationx-property-publisher.md)** задать абсолютные вращения фигуры относительно оси x.

Нельзя изменять размеры поворот вокруг оси x указанного фигуры за границу верхней или нижней, для свойства **RotationX** (90 градусов - 90 градусов). К примеру Если свойство **RotationX** изначально установлено значение 80 и можно указать 40 для аргумента **_порядкового номера_** , итоговый вращение 90 (верхний предел для свойства **RotationX** ) вместо 120.

Чтобы изменить вращения фигуры относительно оси y (по вертикали), используйте метод **[IncrementRotationY](threedformat-incrementrotationy-method-publisher.md)** . Чтобы изменить вращения вокруг оси z (расширяет наружу плоскости публикации), используйте метод **[IncrementRotation](shape-incrementrotation-method-publisher.md)** .


## <a name="example"></a>Пример

В этом примере tilts первую фигуру в активной публикации копирование 10 градусов. Фигуры должен быть вытянутый фигуру увидеть результат этого кода.


```vb
ActiveDocument.Pages(1).Shapes(1).ThreeD _ 
 .IncrementRotationX Increment:=10
```


