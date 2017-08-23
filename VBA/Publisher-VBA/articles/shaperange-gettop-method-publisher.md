---
title: "Метод ShapeRange.GetTop (издатель)"
keywords: vbapb10.chm2293783
f1_keywords: vbapb10.chm2293783
ms.prod: publisher
api_name: Publisher.ShapeRange.GetTop
ms.assetid: bbee5dec-78fd-efd9-1368-2089a44d9bff
ms.date: 06/08/2017
ms.openlocfilehash: 2e219e2b5e13fe93bad5efb4fc0b8fff2cda303d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangegettop-method-publisher"></a>Метод ShapeRange.GetTop (издатель)

Возвращает расстояние диапазона фигуры или фигуры верхнего края из верхнего края самые левые страницы в текущем представлении в виде **одного** в указанных единицах.


## <a name="syntax"></a>Синтаксис

 _выражение_. **GetTop** ( **_Единицы_**)

 переменная _expression_A, представляющий объект **ShapeRange** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Подразделения|Обязательное свойство.| **PbUnitType**|Единицы измерения, в которой требуется получить расстояние.|

### <a name="return-value"></a>Возвращаемое значение

Один


## <a name="remarks"></a>Заметки

Параметр устройства может иметь одно из **[PbUnitType](pbunittype-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.

Метод **[GetLeft](shape-getleft-method-publisher.md)** используется для возврата расстояние фигуры или диапазона фигуры левого края от левого края самые левые страницы в текущем представлении.


## <a name="example"></a>Пример

В следующем примере показан расстояния слева и верхнему краю самые левые страницы слева и верхнего края диапазона фигуры, состоящий из всех фигур на первой странице. Расстояния выражаются в дюймах (до сотой).


```vb
With ActiveDocument.Pages(1).Shapes.Range 
 MsgBox "Distance from left: " _ 
 &; Format(.GetLeft(Unit:=pbUnitInch), "0.00") _ 
 &; " in" &; vbCr _ 
 &; "Distance from top: " _ 
 &; Format(.GetTop(Unit:=pbUnitInch), "0.00") _ 
 &; " in" 
End With
```


