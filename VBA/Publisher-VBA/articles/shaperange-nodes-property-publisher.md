---
title: "Свойство ShapeRange.Nodes (издатель)"
keywords: vbapb10.chm2293829
f1_keywords: vbapb10.chm2293829
ms.prod: publisher
api_name: Publisher.ShapeRange.Nodes
ms.assetid: 513be66c-558c-f5f3-ed89-0ef4bc5a0101
ms.date: 06/08/2017
ms.openlocfilehash: 059beac475fc1633e67bd4ceb299b31b8d4a102f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangenodes-property-publisher"></a>Свойство ShapeRange.Nodes (издатель)

Возвращает коллекцию **[ShapeNodes](shapenodes-object-publisher.md)** , представляющий геометрическое описание указанной фигуры. Применяется к **фигуры** или **ShapeRange** объектов, представляющих freeform документы.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Узлы**

 переменная _expression_A, представляющий объект **ShapeRange** .


## <a name="example"></a>Пример

В этом примере добавляется легко узел с сегмент после узла четырех в три фигуры на одну страницу. Фигура трех должен быть freeform документа по крайней мере четыре узлами.


```vb
With ActiveDocument.Pages(1) _ 
 .Shapes(3).Nodes 
 .Insert Index:=4, SegmentType:=msoSegmentCurve, _ 
 EditingType:=msoEditingSmooth, X1:=210, Y1:=100 
End With
```


