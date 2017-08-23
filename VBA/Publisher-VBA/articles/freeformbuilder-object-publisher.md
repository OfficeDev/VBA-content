---
title: "Объект FreeformBuilder (издатель)"
keywords: vbapb10.chm3342335
f1_keywords: vbapb10.chm3342335
ms.prod: publisher
api_name: Publisher.FreeformBuilder
ms.assetid: 542df9f7-f636-a98e-01de-11005b5797cc
ms.date: 06/08/2017
ms.openlocfilehash: 2741714df93e14e4bafe781ee40c5f3facda162f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="freeformbuilder-object-publisher"></a>Объект FreeformBuilder (издатель)

Представляет геометрии freeform во время его построения.
 


## <a name="example"></a>Пример

Метод **[BuildFreeform](shapes-buildfreeform-method-publisher.md)** коллекцию **[фигур](shapes-object-publisher.md)** для возврата объекта **FreeformBuilder** . Чтобы добавить узлы в фигуру, используйте метод **[AddNodes](freeformbuilder-addnodes-method-publisher.md)** . Используйте метод **[ConvertToShape](freeformbuilder-converttoshape-method-publisher.md)** для создания фигуры, определенные в объекте **FreeformBuilder** и добавить его в коллекцию **фигур** . В следующем примере добавляется freeform с четырьмя сегменты в активный документ.
 

 

```
Sub CreateNewFreeFormShape() 
 With ActiveDocument.Pages(1).Shapes.BuildFreeform( _ 
 EditingType:=msoEditingCorner, X1:=360, Y1:=200) 
 .AddNodes SegmentType:=msoSegmentCurve, _ 
 EditingType:=msoEditingCorner, X1:=380, Y1:=230, _ 
 X2:=400, Y2:=250, X3:=450, Y3:=300 
 .AddNodes SegmentType:=msoSegmentCurve, _ 
 EditingType:=msoEditingAuto, X1:=480, Y1:=200 
 .AddNodes SegmentType:=msoSegmentLine, _ 
 EditingType:=msoEditingAuto, X1:=480, Y1:=400 
 .AddNodes SegmentType:=msoSegmentLine, _ 
 EditingType:=msoEditingAuto, X1:=360, Y1:=200 
 .ConvertToShape 
 End With 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[AddNodes](freeformbuilder-addnodes-method-publisher.md)|
|[ConvertToShape](freeformbuilder-converttoshape-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](freeformbuilder-application-property-publisher.md)|
|[Родительский раздел](freeformbuilder-parent-property-publisher.md)|

