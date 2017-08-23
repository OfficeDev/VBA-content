---
title: "Свойство Options.MeasurementUnit (издатель)"
keywords: vbapb10.chm1048594
f1_keywords: vbapb10.chm1048594
ms.prod: publisher
api_name: Publisher.Options.MeasurementUnit
ms.assetid: 49221e4e-c84a-6706-8f9a-3853283ebb18
ms.date: 06/08/2017
ms.openlocfilehash: 4986b9aa615e2ee528f29e113ab287d6a5249567
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="optionsmeasurementunit-property-publisher"></a>Свойство Options.MeasurementUnit (издатель)

Возвращает или задает значение константы **PbUnitType** , представляющее единицы измерения standard для Microsoft Publisher. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MeasurementUnit**

 переменная _expression_A, представляет собой объект- **Параметры** .


### <a name="return-value"></a>Возвращаемое значение

PbUnitType


## <a name="remarks"></a>Заметки

Значение свойства **MeasurementUnit** может иметь одно из **PbUnitType** константы объявляются в библиотеке типов издателя и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **pbUnitCM**| Задает единицы измерения см.|
| **pbUnitEmu**| Не определено для этого свойства. Возвращает ошибку, если используется...|
| **pbUnitFeet**|Не определено для этого свойства. Возвращает ошибку, если используется.|
| **pbUnitHa**|Не определено для этого свойства. Возвращает ошибку, если используется.|
| **pbUnitInch**|Задает единицы измерения см.|
| **pbUnitKyu**| Не определено для этого свойства. Возвращает ошибку, если используется.|
| **pbUnitMeter** .|Не определено для этого свойства. Возвращает ошибку, если используется.|
| **pbUnitPica**|Задает единицы измерения пики.|
| **pbUnitPoint**|Задает единицы измерения в точках.|
| **pbUnitTwip**|Не определено для этого свойства. Возвращает ошибку, если используется.|

## <a name="example"></a>Пример

В этом примере задается единицы измерения standard для Publisher точек.


```vb
Sub SetUnitOfMeasurement() 
 Options.MeasurementUnit = pbUnitPoint 
End Sub
```

В этом примере отображается текущий единицы измерения.




```vb
Sub GetUnitOfMeasurement() 
 Dim measUnit As PbUnitType 
 Dim strUnit As String 
 
 measUnit = Options.MeasurementUnit 
 
 Select Case measUnit 
 Case 0 
 strUnit = "inches" 
 Case 1 
 strUnit = "centimeters" 
 Case 2 
 strUnit = "picas" 
 Case 3 
 strUnit = "points" 
 End Select 
 
 MsgBox "The current unit of measurement is " &; strUnit 
 
End Sub
```


