---
title: "Метод ShadowFormat.IncrementOffsetY (издатель)"
keywords: vbapb10.chm3670033
f1_keywords: vbapb10.chm3670033
ms.prod: publisher
api_name: Publisher.ShadowFormat.IncrementOffsetY
ms.assetid: fca7a688-adf8-d8cd-8e14-9d1988c8d9f2
ms.date: 06/08/2017
ms.openlocfilehash: 2d29e62588f9e576c54ea41cee5d58651edbfe45
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shadowformatincrementoffsety-method-publisher"></a>Метод ShadowFormat.IncrementOffsetY (издатель)

Постепенно изменяется вертикальное смещение тени на определенное расстояние.


## <a name="syntax"></a>Синтаксис

 _выражение_. **IncrementOffsetY** ( **_Порядкового номера_**)

 переменная _expression_A, представляет собой объект- **ShadowFormat** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Порядкового номера|Обязательное свойство.| **Variant**|Указывает, насколько смещение тени перемещаемых по вертикали. Положительное значение перемещает тени; отрицательное значение перемещает его вверх. Числовые значения вычисляются в точках; строк может быть в любой устройств, поддерживаемых Microsoft Publisher (например, «2,5 дюйма»).|

## <a name="remarks"></a>Заметки

Свойство **[OffsetY](shadowformat-offsety-property-publisher.md)** задать смещение абсолютный вертикальной тени.

Используйте метод **[IncrementOffsetX](shadowformat-incrementoffsetx-method-publisher.md)** для изменения горизонтальной смещение тени.


## <a name="example"></a>Пример

В этом примере Сдвиг тени для третьего фигуры в активной публикации 3 точки.


```vb
ActiveDocument.Pages(1).Shapes(3).Shadow _ 
 .IncrementOffsetY Increment:=-3 

```


