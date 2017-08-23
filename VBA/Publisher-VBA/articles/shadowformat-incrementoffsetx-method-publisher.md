---
title: "Метод ShadowFormat.IncrementOffsetX (издатель)"
keywords: vbapb10.chm3670032
f1_keywords: vbapb10.chm3670032
ms.prod: publisher
api_name: Publisher.ShadowFormat.IncrementOffsetX
ms.assetid: 05c25f0f-beac-2b25-630b-57d4a3bdb0c9
ms.date: 06/08/2017
ms.openlocfilehash: fed9a8c966629b2f6111560194f81bb5ff6e3e0b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shadowformatincrementoffsetx-method-publisher"></a>Метод ShadowFormat.IncrementOffsetX (издатель)

Постепенно меняет горизонтальное смещение тени на определенное расстояние.


## <a name="syntax"></a>Синтаксис

 _выражение_. **IncrementOffsetX** ( **_Порядкового номера_**)

 переменная _expression_A, представляет собой объект- **ShadowFormat** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Порядкового номера|Обязательное свойство.| **Variant**|Указывает, насколько смещение тени перемещаемых по горизонтали. Положительное значение перемещает тени вправо; отрицательное значение перемещает его слева. Числовые значения вычисляются в точках; строк может быть в любой устройств, поддерживаемых Microsoft Publisher (например, «2,5 дюйма»).|

## <a name="remarks"></a>Заметки

Свойство **[OffsetX](shadowformat-offsetx-property-publisher.md)** задать смещение абсолютный горизонтальной тени.

Используйте метод **[IncrementOffsetY](shadowformat-incrementoffsety-method-publisher.md)** для изменения вертикальной смещение тени.


## <a name="example"></a>Пример

В этом примере Сдвиг тени для третьего фигуры в активной публикации слева 3 точки.


```vb
ActiveDocument.Pages(1).Shapes(3).Shadow _ 
 .IncrementOffsetX Increment:=-3 

```


