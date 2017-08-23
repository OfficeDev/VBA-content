---
title: "Метод Shape.IncrementLeft (издатель)"
keywords: vbapb10.chm2228256
f1_keywords: vbapb10.chm2228256
ms.prod: publisher
api_name: Publisher.Shape.IncrementLeft
ms.assetid: 447886ad-f515-9869-524a-a803ab025fa4
ms.date: 06/08/2017
ms.openlocfilehash: 13e9122f938f9c9354b4f0963f251038765b6b12
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapeincrementleft-method-publisher"></a>Метод Shape.IncrementLeft (издатель)

Перемещает указанные форму или диапазона фигуры по горизонтали на определенное расстояние.


## <a name="syntax"></a>Синтаксис

 _выражение_. **IncrementLeft** ( **_Порядкового номера_**)

 переменная _expression_A, представляющий объект **фигуры** .


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


