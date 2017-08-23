---
title: "Объект ShapeNodes (издатель)"
keywords: vbapb10.chm3538943
f1_keywords: vbapb10.chm3538943
ms.prod: publisher
api_name: Publisher.ShapeNodes
ms.assetid: f190a8a8-e03a-e8a2-482a-5e092ff3ed86
ms.date: 06/08/2017
ms.openlocfilehash: af3f5ab4310af0344eab03b20c36adfc34546463
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapenodes-object-publisher"></a>Объект ShapeNodes (издатель)

Коллекция всех **[ShapeNode](shapenode-object-publisher.md)** объектов в указанном freeform. Каждый объект **ShapeNode** представляет узел между сегменты в произвольный или контрольной точки для сегмент произвольной формы. Можно создать freeform вручную или с помощью методов **[BuildFreeform](shapes-buildfreeform-method-publisher.md)** и **[ConvertToShape](freeformbuilder-converttoshape-method-publisher.md)** .
 


## <a name="example"></a>Пример

Используйте свойство **[узлов](shape-nodes-property-publisher.md)** для возврата коллекции **ShapeNodes** . В следующем примере удаляется узел четырех в трех фигуры на активном документе. В данном примере для работы фигуры три значения freeform по крайней мере четыре узлами.
 

 

```
Sub DeleteShapeNode() 
 ActiveDocument.Pages(1).Shapes(3).Nodes.Delete Index:=4 
End Sub
```

Используйте метод **[Insert](shapenodes-insert-method-publisher.md)** для создания нового узла и добавления его в коллекцию **ShapeNodes** . Следующий пример добавляет легко узел с сегмент после узла четырех в трех фигуры на активном документе. В данном примере для работы фигуры три значения e freeform по крайней мере четыре узлами.
 

 



```
Sub AddCurvedSmoothSegment() 
 ActiveDocument.Pages(1).Shapes(3).Nodes.Insert _ 
 Index:=4, SegmentType:=msoSegmentCurve, _ 
 EditingType:=msoEditingSmooth, X1:=210, Y1:=100 
End Sub
```

С помощью **узлов** (индекс), где индекс — номер индекса узла, для возврата объекта **ShapeNode** . Если один узел в трех фигуры на активном документе точку угла, следующий пример делает точку смягчения. В данном примере для работы фигуры три значения freeform.
 

 



```
Sub SetPointType() 
 With ActiveDocument.Pages(1).Shapes(3) 
 If .Nodes(1).EditingType = msoEditingCorner Then 
 .Nodes.SetEditingType Index:=1, EditingType:=msoEditingSmooth 
 End If 
 End With 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Delete](shapenodes-delete-method-publisher.md)|
|[Вставка](shapenodes-insert-method-publisher.md)|
|[Элемент](shapenodes-item-method-publisher.md)|
|[SetEditingType](shapenodes-seteditingtype-method-publisher.md)|
|[SetPosition](shapenodes-setposition-method-publisher.md)|
|[SetSegmentType](shapenodes-setsegmenttype-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](shapenodes-application-property-publisher.md)|
|[Count](shapenodes-count-property-publisher.md)|
|[Родительский раздел](shapenodes-parent-property-publisher.md)|

