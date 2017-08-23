---
title: "Свойство Shape.Nodes (издатель)"
keywords: vbapb10.chm2228293
f1_keywords: vbapb10.chm2228293
ms.prod: publisher
api_name: Publisher.Shape.Nodes
ms.assetid: a1463ff3-5b75-e4b9-df12-985538713c7c
ms.date: 06/08/2017
ms.openlocfilehash: 17262b3aba8a701f8cd47e8a7f3d7a604d42f123
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapenodes-property-publisher"></a>Свойство Shape.Nodes (издатель)

Возвращает коллекцию **[ShapeNodes](shapenodes-object-publisher.md)** , представляющий геометрическое описание указанной фигуры. Применяется к **фигуры** или **ShapeRange** объектов, представляющих freeform документы.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Узлы**

 переменная _expression_A, представляющий объект **фигуры** .


## <a name="example"></a>Пример

В этом примере добавляется легко узел с сегмент после узла четырех в три фигуры на одну страницу. Фигура трех должен быть freeform документа по крайней мере четыре узлами.


```vb
With ActiveDocument.Pages(1) _ 
 .Shapes(3).Nodes 
 .Insert Index:=4, SegmentType:=msoSegmentCurve, _ 
 EditingType:=msoEditingSmooth, X1:=210, Y1:=100 
End With
```


