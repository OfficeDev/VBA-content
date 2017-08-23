---
title: "Метод Shape.IncrementTop (издатель)"
keywords: vbapb10.chm2228258
f1_keywords: vbapb10.chm2228258
ms.prod: publisher
api_name: Publisher.Shape.IncrementTop
ms.assetid: c7a5bf47-7c5a-f6e8-b2b7-c95bea9dc081
ms.date: 06/08/2017
ms.openlocfilehash: 9e0dc0286041ba50d113784cb964cdbf46a3db84
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapeincrementtop-method-publisher"></a>Метод Shape.IncrementTop (издатель)

Перемещает указанные форму или диапазона фигуры на определенное расстояние по вертикали.


## <a name="syntax"></a>Синтаксис

 _выражение_. **IncrementTop** ( **_Порядкового номера_**)

 переменная _expression_A, представляющий объект **фигуры** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Порядкового номера|Обязательное свойство.| **Variant**|Расстояние по вертикали для перемещения фигуры или диапазона фигуры. Положительное значение перемещает форму или диапазона фигуры вниз; отрицательное значение перемещает его вверх. Числовые значения вычисляются в точках; строк может быть в любой устройств, поддерживаемых Microsoft Publisher (например, «2,5 дюйма»).|

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


