---
title: "Метод ShapeNodes.SetEditingType (издатель)"
keywords: vbapb10.chm3473427
f1_keywords: vbapb10.chm3473427
ms.prod: publisher
api_name: Publisher.ShapeNodes.SetEditingType
ms.assetid: f90b1323-d682-1b2b-6747-cea5f2cead3c
ms.date: 06/08/2017
ms.openlocfilehash: 9223ba290e35e0a4cca0a40371ab22077a009d0f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapenodesseteditingtype-method-publisher"></a>Метод ShapeNodes.SetEditingType (издатель)

Задает тип редактирования указанного узла. Если узел является контрольной точки для сегмент, этого метода можно задать редактирования тип узла рядом с ней, соединяет два сегмента. В зависимости от типа редактирования этот метод может повлиять на положение рядом с узлами.


## <a name="syntax"></a>Синтаксис

 _выражение_. **SetEditingType** ( **_Индекса_**, **_EditingType_**)

 переменная _expression_A, представляет собой объект- **ShapeNodes** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Индекс|Обязательное свойство.| **Длинный**|Узел, тип которого редактирования не должно быть задано. Должно быть число от 1 до количества узлов в указанном фигуры; в противном случае возникает ошибка.|
|EditingType|Обязательное свойство.| **MsoEditingType**|Свойство редактирования узла.|

## <a name="remarks"></a>Заметки

Параметр EditingType может быть одной из констант **MsoEditingType** объявлена в библиотеке типов, Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoEditingAuto**|Изменяет узел типа подходят для сегменты подключения.|
| **msoEditingCorner**| Изменяет узел узел угла.|
| **msoEditingSmooth**|Изменение узла на узел легко график...|
| **msoEditingSymmetric**|Изменяет узел симметричного график узел.|

## <a name="example"></a>Пример

В этом примере изменяется все узлы углу сгладить узлов в третьей фигуры active публикации. Фигуры должен быть freeform документа.


```vb
Dim intNode As Integer 
 
With ActiveDocument.Pages(1).Shapes(3).Nodes 
 For intNode = 1 to .Count 
 If .Item(intNode).EditingType = msoEditingCorner Then 
 .SetEditingType _ 
 Index:=intNode, EditingType:=msoEditingSmooth 
 End If 
 Next intNode 
End With 

```


