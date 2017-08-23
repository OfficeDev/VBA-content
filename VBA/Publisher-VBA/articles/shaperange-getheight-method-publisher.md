---
title: "Метод ShapeRange.GetHeight (издатель)"
keywords: vbapb10.chm2293784
f1_keywords: vbapb10.chm2293784
ms.prod: publisher
api_name: Publisher.ShapeRange.GetHeight
ms.assetid: 63501bf7-c24d-b58e-e4c5-c8a229f07c4e
ms.date: 06/08/2017
ms.openlocfilehash: dfb3882c6f3a89c88d16772135b04c5c994762df
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangegetheight-method-publisher"></a>Метод ShapeRange.GetHeight (издатель)

Возвращает высоту фигуры или диапазона фигуры в виде **одного** в указанных единицах.


## <a name="syntax"></a>Синтаксис

 _выражение_. **GetHeight** ( **_Единицы_**)

 переменная _expression_A, представляющий объект **ShapeRange** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Подразделения|Обязательное свойство.| **PbUnitType**|Единицы измерения, в которой требуется получить высоту.|

### <a name="return-value"></a>Возвращаемое значение

Один


## <a name="remarks"></a>Заметки

Параметр устройства может иметь одно из **[PbUnitType](pbunittype-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.

Используйте метод **[GetWidth](shape-getwidth-method-publisher.md)** для возврата ширины фигуры или диапазона фигуры.


## <a name="example"></a>Пример

Следующий пример отображает высоту и ширину в дюймах (до сотой) диапазона фигуры, состоящий из всех фигур на первой странице active публикации.


```vb
With ActiveDocument.Pages(1).Shapes.Range 
 MsgBox "Height of all shapes: " _ 
 &; Format(.GetHeight(Unit:=pbUnitInch), "0.00") _ 
 &; " in" &; vbCr _ 
 &; "Width of all shapes: " _ 
 &; Format(.GetWidth(Unit:=pbUnitInch), "0.00") _ 
 &; " in" 
End With 

```


