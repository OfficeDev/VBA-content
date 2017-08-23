---
title: "Объект ShapeNode (издатель)"
keywords: vbapb10.chm3604479
f1_keywords: vbapb10.chm3604479
ms.prod: publisher
api_name: Publisher.ShapeNode
ms.assetid: 8246e1fd-2477-91f4-490b-2d2b6032fccd
ms.date: 06/08/2017
ms.openlocfilehash: 28a657f43997fea5b556c80ce9dfeb7e6dd31f18
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapenode-object-publisher"></a>Объект ShapeNode (издатель)

Представляет геометрии и редактирования геометрии свойства узлов в определенный пользователем freeform. Узлы включают грани между сегменты фигуру и контрольные точки для изогнутые сегменты. Объект **ShapeNode** является элементом коллекции **[ShapeNodes](shapenodes-object-publisher.md)** . Коллекция **ShapeNodes** содержит все узлы в произвольной формы.
 


## <a name="example"></a>Пример

С помощью **узлов** (индекс), где индекс — номер индекса узла, для возврата объекта **ShapeNode** . Если один узел в трех фигуры на активном документе точку угла, следующий пример делает точку смягчения. В данном примере для работы фигуры один должен быть freeform.
 

 

```
Sub ChangeNodeType() 
 With ActiveDocument.Pages(1).Shapes(1) 
 If .Nodes(1).EditingType = msoEditingCorner Then 
 .Nodes.SetEditingType Index:=1, EditingType:=msoEditingSmooth 
 End If 
 End With 
End Sub
```


## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](shapenode-application-property-publisher.md)|
|[EditingType](shapenode-editingtype-property-publisher.md)|
|[Родительский раздел](shapenode-parent-property-publisher.md)|
|[Точки](shapenode-points-property-publisher.md)|
|[SegmentType](shapenode-segmenttype-property-publisher.md)|

