---
title: "Метод ShapeRange.IncrementTop (издатель)"
keywords: vbapb10.chm2293794
f1_keywords: vbapb10.chm2293794
ms.prod: publisher
api_name: Publisher.ShapeRange.IncrementTop
ms.assetid: 8172406f-fac5-ad3d-49b8-cb4858d45c6d
ms.date: 06/08/2017
ms.openlocfilehash: 98c8966c25f66e94427fb05d46244f836c69ed91
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangeincrementtop-method-publisher"></a>Метод ShapeRange.IncrementTop (издатель)

Перемещает указанные форму или диапазона фигуры на определенное расстояние по вертикали.


## <a name="syntax"></a>Синтаксис

 _выражение_. **IncrementTop** ( **_Порядкового номера_**)

 переменная _expression_A, представляющий объект **ShapeRange** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Порядкового номера|Обязательное свойство.| **Variant**|Расстояние по вертикали для перемещения фигуры или диапазона фигуры. Положительное значение перемещает форму или диапазона фигуры вниз; отрицательное значение перемещает его вверх. Числовые значения вычисляются в точках; строк может быть в любой устройств, поддерживаемых Microsoft Publisher (например, «2,5 дюйма»).|

### <a name="return-value"></a>Возвращаемое значение

Значение Nothing


## <a name="remarks"></a>Заметки

Используйте метод **[IncrementLeft](shape-incrementleft-method-publisher.md)** для перемещения фигур или диапазоны фигуры по горизонтали.


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


