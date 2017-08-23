---
title: "Метод ShapeRange.IncrementLeft (издатель)"
keywords: vbapb10.chm2293792
f1_keywords: vbapb10.chm2293792
ms.prod: publisher
api_name: Publisher.ShapeRange.IncrementLeft
ms.assetid: 1b760b5d-9879-5f64-c4c5-c9834a7928ff
ms.date: 06/08/2017
ms.openlocfilehash: 59e16a1612d5db3cf46c8e222ff00d9bfad80216
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangeincrementleft-method-publisher"></a>Метод ShapeRange.IncrementLeft (издатель)

Перемещает указанные форму или диапазона фигуры по горизонтали на определенное расстояние.


## <a name="syntax"></a>Синтаксис

 _выражение_. **IncrementLeft** ( **_Порядкового номера_**)

 переменная _expression_A, представляющий объект **ShapeRange** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Порядкового номера|Обязательное свойство.| **Variant**|Расстояние по горизонтали переместить форму или диапазона фигуры. Положительное значение перемещает форму или диапазона фигуры вправо; отрицательное значение перемещает его слева. Числовые значения вычисляются в точках; строк может быть в любой устройств, поддерживаемых Microsoft Publisher (например, «2,5 дюйма»).|

### <a name="return-value"></a>Возвращаемое значение

Значение Nothing


## <a name="remarks"></a>Заметки

Используйте метод **[IncrementTop](shape-incrementtop-method-publisher.md)** для перемещения фигур или диапазоны фигуры по вертикали.


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


