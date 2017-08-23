---
title: "Метод Shape.IncrementRotation (издатель)"
keywords: vbapb10.chm2228257
f1_keywords: vbapb10.chm2228257
ms.prod: publisher
api_name: Publisher.Shape.IncrementRotation
ms.assetid: 3293c707-f3e8-1afb-cf9c-231ceae66ab6
ms.date: 06/08/2017
ms.openlocfilehash: 483cd856873c8d96e980eb80ad5e3fdcb87eb540
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapeincrementrotation-method-publisher"></a>Метод Shape.IncrementRotation (издатель)

Изменяет вращение указанного фигуры относительно оси z (расширяет наружу плоскости публикации) указанное число градусов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **IncrementRotation** ( **_Порядкового номера_**)

 переменная _expression_A, представляющий объект **фигуры** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Порядкового номера|Обязательное свойство.| **Один**|Указывает, насколько фигуры вращаться вокруг оси z в градусов. Положительное значение Поворот фигуры часовой; отрицательное значение поворот против. Допустимые значения: от - 360 до 360.|

### <a name="return-value"></a>Возвращаемое значение

Значение Nothing


## <a name="remarks"></a>Заметки

Свойство **[Вращение](shaperange-rotation-property-publisher.md)** задать абсолютные Поворот фигуры.

Поворот объемной фигуры вокруг оси x (по горизонтали) или y (по вертикали), используйте метод **[IncrementRotationX](threedformat-incrementrotationx-method-publisher.md)** или **[IncrementRotationY](threedformat-incrementrotationy-method-publisher.md)** , соответственно.


## <a name="example"></a>Пример

В этом примере дубликатов первую фигуру на активной публикации задает заливки для повторяющихся, перемещает 70 точек вправо и на 50 точек вверх и поворот его 30 градусов часовой.


```vb
With ActiveDocument.Pages(1).Shapes(1).Duplicate 
 .Fill.PresetTextured PresetTexture:=msoTextureGranite 
 .IncrementLeft Increment:=70 
 .IncrementTop Increment:=-50 
 .IncrementRotation Increment:=30 
End With
```


