---
title: "Метод Shape.GetLeft (издатель)"
keywords: vbapb10.chm2228246
f1_keywords: vbapb10.chm2228246
ms.prod: publisher
api_name: Publisher.Shape.GetLeft
ms.assetid: e8f28ab3-f9da-eae7-2a21-b8b2505e9b44
ms.date: 06/08/2017
ms.openlocfilehash: dc9b760ab4dd8846fcc3b23e649a4cd674f0744f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapegetleft-method-publisher"></a>Метод Shape.GetLeft (издатель)

Возвращает расстояние от левого края диапазона фигуры или фигуры от левого края самые левые страницы в текущем представлении в виде **одного** в указанных единицах.


## <a name="syntax"></a>Синтаксис

 _выражение_. **GetLeft** ( **_Единицы_**)

 переменная _expression_A, представляющий объект **фигуры** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Подразделения|Обязательное свойство.| **PbUnitType**|Единицы измерения, в которой требуется получить расстояние.|

### <a name="return-value"></a>Возвращаемое значение

Один


## <a name="remarks"></a>Заметки

Параметр устройства может иметь одно из **[PbUnitType](pbunittype-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.

Используйте метод **[GetTop](shape-gettop-method-publisher.md)** для возврата расстояние диапазона фигуры или фигуры верхнего края из верхнего края самые левые страницы в текущем представлении.


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


