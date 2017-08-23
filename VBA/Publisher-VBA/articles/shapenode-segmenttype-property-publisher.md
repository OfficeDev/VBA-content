---
title: "Свойство ShapeNode.SegmentType (издатель)"
keywords: vbapb10.chm3539202
f1_keywords: vbapb10.chm3539202
ms.prod: publisher
api_name: Publisher.ShapeNode.SegmentType
ms.assetid: 471206b2-ca37-5e4a-678b-df8a47c90f96
ms.date: 06/08/2017
ms.openlocfilehash: 130ec238af067924fc3dbb67bc666c15f7e05d97
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapenodesegmenttype-property-publisher"></a>Свойство ShapeNode.SegmentType (издатель)

Возвращает константу **MsoSegmentType** , указывающее прямых или изогнутых сегмента, связанного с указанного узла. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **SegmentType**

 переменная _expression_A, представляет собой объект- **ShapeNode** .


### <a name="return-value"></a>Возвращаемое значение

MsoSegmentType


## <a name="remarks"></a>Заметки

Если указанный узел контрольной точки для сегмент, данное свойство возвращает **msoSegmentCurve**.

Используйте метод **[SetSegmentType](shapenodes-setsegmenttype-method-publisher.md)** для задания значения этого свойства.

Значение свойства **SegmentType** может иметь одно из следующих **MsoSegmentType** константы, описанные в библиотеке типов, Microsoft Publisher.



| **msoSegmentCurve**|| **msoSegmentLine**|

## <a name="example"></a>Пример

В этом примере изменяется все прямые сегменты изогнутые сегменты в первую фигуру на первой странице active публикации. В данном примере для работы указанного фигуры должен быть freeform документа.


```vb
Sub ChangeSegmentTypes() 
 Dim intNode As Integer 
 With ActiveDocument.Pages(1).Shapes(1).Nodes 
 intNode = 1 
 Do While intNode <= .Count 
 If .Item(intNode).SegmentType = msoSegmentLine Then 
 .SetSegmentType Index:=intNode, _ 
 SegmentType:=msoSegmentCurve 
 End If 
 intNode = intNode + 1 
 Loop 
 End With 
End Sub
```


