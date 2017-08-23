---
title: "Метод FreeformBuilder.ConvertToShape (издатель)"
keywords: vbapb10.chm3276817
f1_keywords: vbapb10.chm3276817
ms.prod: publisher
api_name: Publisher.FreeformBuilder.ConvertToShape
ms.assetid: 1cb490af-40be-b03f-2f8d-04b1015fbde3
ms.date: 06/08/2017
ms.openlocfilehash: a8b15e17187513aef24558f04c067024f719945f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="freeformbuilderconverttoshape-method-publisher"></a>Метод FreeformBuilder.ConvertToShape (издатель)

Создает фигуры, геометрические характеристики на указанный объект **[FreeformBuilder](freeformbuilder-object-publisher.md)** . Возвращает объект **[фигуры](shape-object-publisher.md)** , представляющий новую фигуру.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ConvertToShape**

 переменная _expression_A, представляет собой объект- **FreeformBuilder** .


### <a name="return-value"></a>Возвращаемое значение

Shape


## <a name="remarks"></a>Заметки

Необходимо применить метод **[AddNodes](freeformbuilder-addnodes-method-publisher.md)** объекта **FreeformBuilder** по крайней мере один раз перед используйте метод **ConvertToShape** или возникает ошибка.


## <a name="example"></a>Пример

В этом примере добавляется freeform с четырьмя вершинами для первой страницы в активной публикации.


```vb
' Add a new freeform object. 
With ActiveDocument.Shapes _ 
 .BuildFreeform(EditingType:=msoEditingCorner, _ 
 X1:=100, Y1:=100) 
 
 ' Add three more nodes and close the polygon. 
 .AddNodes SegmentType:=msoSegmentCurve, _ 
 EditingType:=msoEditingCorner, _ 
 X1:=200, Y1:=200, X2:=225, Y2:=250, X3:=250, Y3:=200 
 .AddNodes SegmentType:=msoSegmentCurve, _ 
 EditingType:=msoEditingAuto, X1:=200, Y1:=100 
 .AddNodes SegmentType:=msoSegmentLine, _ 
 EditingType:=msoEditingAuto, X1:=150, Y1:=50 
 .AddNodes SegmentType:=msoSegmentLine, _ 
 EditingType:=msoEditingAuto, X1:=100, Y1:=100 
 
 ' Convert the polygon to a Shape object. 
 .ConvertToShape 
End With 

```


