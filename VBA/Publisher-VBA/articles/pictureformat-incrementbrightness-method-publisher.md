---
title: "Метод PictureFormat.IncrementBrightness (издатель)"
keywords: vbapb10.chm3604496
f1_keywords: vbapb10.chm3604496
ms.prod: publisher
api_name: Publisher.PictureFormat.IncrementBrightness
ms.assetid: 912fd08e-bbb3-bf98-b0da-7128926f3409
ms.date: 06/08/2017
ms.openlocfilehash: 18575d73fb991936954e875061c4139eb2c80e80
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformatincrementbrightness-method-publisher"></a>Метод PictureFormat.IncrementBrightness (издатель)

Изменение яркости рисунка на указанную величину.


## <a name="syntax"></a>Синтаксис

 _выражение_. **IncrementBrightness** ( **_Порядкового номера_**)

 переменная _expression_A, представляет собой объект- **PictureFormat** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Порядкового номера|Обязательное свойство.| **Один**|Определяет, насколько изменение значение свойства **[яркость](pictureformat-brightness-property-publisher.md)** рисунка. Положительное значение делает изображение более яркие; отрицательное значение делает изображение более темные. Допустимые значения: от - 1 до 1.|

## <a name="remarks"></a>Заметки

Не удается яркость изображения за границу верхней или нижней, для свойства **яркость** . К примеру Если свойство **яркость** изначально установлено значение 0,9 и можно указать 0,3 для аргумента **_порядкового номера_** , итоговый уровень яркости 1.0, являющийся верхний предел для свойства **яркость** , вместо 1.2.

Свойство **яркость** задать абсолютные яркости рисунка.


## <a name="example"></a>Пример

В этом примере создается дубликат первой фигуры в активной публикации и перемещает и затемняет дубликата. Для обеспечения работы примера фигуры значения рисунок или объект OLE, представляющий изображение.


```vb
With ActiveDocument.Pages(1).Shapes(1).Duplicate 
 .PictureFormat.IncrementBrightness Increment:=-0.2 
 .IncrementLeft Increment:=50 
 .IncrementTop Increment:=50 
End With 

```


