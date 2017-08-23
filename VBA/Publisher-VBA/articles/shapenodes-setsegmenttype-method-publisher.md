---
title: "Метод ShapeNodes.SetSegmentType (издатель)"
keywords: vbapb10.chm3473429
f1_keywords: vbapb10.chm3473429
ms.prod: publisher
api_name: Publisher.ShapeNodes.SetSegmentType
ms.assetid: 64f742fb-8216-9ec3-3fa9-ca2b319cf3e9
ms.date: 06/08/2017
ms.openlocfilehash: 4f937c2ae289730fce20ea7a6e4409e2c5ec30f2
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapenodessetsegmenttype-method-publisher"></a>Метод ShapeNodes.SetSegmentType (издатель)

Задает тип сегмента сегмент, исходя из указанного узла. Если узел является контрольной точки для сегмент, этого метода можно задать тип сегмента для этого график; Это может повлиять на общее число узлов Вставка или удаление рядом с узлами.


## <a name="syntax"></a>Синтаксис

 _выражение_. **SetSegmentType** ( **_Индекса_**, **_SegmentType_**)

 переменная _expression_A, представляет собой объект- **ShapeNodes** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Индекс|Обязательное свойство.| **Длинный**|Узел, тип которого сегмент не должно быть задано. Должно быть число от 1 до количества узлов в указанном фигуры; в противном случае возникает ошибка.|
|SegmentType|Обязательное свойство.| **MsoSegmentType**|Указывает тип сегмента.|

## <a name="remarks"></a>Заметки

Параметр SegmentType может быть одной из констант **MsoSegmentType** объявлена в библиотеке типов, Microsoft Office и показаны в следующей таблице.



| **msoSegmentCurve**|| **msoSegmentLine**|

## <a name="example"></a>Пример

В этом примере изменяется все прямые сегменты изогнутые сегменты в третьей фигуры в активной публикации. Фигуры должен быть freeform документа.


```vb
Dim intCount As Integer 
 
With ActiveDocument.Pages(1).Shapes(3).Nodes 
 intCount = 1 
 Do While intCount <= .Count 
 If .Item(intCount).SegmentType = msoSegmentLine Then 
 .SetSegmentType _ 
 Index:=intCount, SegmentType:=msoSegmentCurve 
 End If 
 intCount = intCount + 1 
 Loop 
End With 

```


