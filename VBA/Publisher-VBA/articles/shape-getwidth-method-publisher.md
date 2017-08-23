---
title: "Метод Shape.GetWidth (издатель)"
keywords: vbapb10.chm2228249
f1_keywords: vbapb10.chm2228249
ms.prod: publisher
api_name: Publisher.Shape.GetWidth
ms.assetid: 9df33329-c37b-82f5-93b4-fc4752ee907e
ms.date: 06/08/2017
ms.openlocfilehash: 4f8c7f1b6b1852d450d5b62638645a61d4bc02fd
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapegetwidth-method-publisher"></a>Метод Shape.GetWidth (издатель)

Возвращает ширину фигуры или диапазона фигуры в виде **одного** в указанных единицах.


## <a name="syntax"></a>Синтаксис

 _выражение_. **GetWidth** ( **_Единицы_**)

 переменная _expression_A, представляющий объект **фигуры** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Подразделения|Обязательное свойство.| **PbUnitType**|Единицы измерения, в которой требуется получить ширины.|

### <a name="return-value"></a>Возвращаемое значение

Один


## <a name="remarks"></a>Заметки

Параметр устройства может иметь одно из **[PbUnitType](pbunittype-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.

Используйте метод **[GetHeight](shape-getheight-method-publisher.md)** для возврата высоту фигуры или диапазона фигуры.


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


