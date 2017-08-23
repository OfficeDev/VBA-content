---
title: "Объект GroupShapes (издатель)"
keywords: vbapb10.chm3407871
f1_keywords: vbapb10.chm3407871
ms.prod: publisher
api_name: Publisher.GroupShapes
ms.assetid: dd723f99-25a9-81cc-1d16-eb7dcd651c5e
ms.date: 06/08/2017
ms.openlocfilehash: 03b98763a8adf545ca924a7dacc3c74153e93028
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="groupshapes-object-publisher"></a>Объект GroupShapes (издатель)

Представляет отдельные фигуры в группы фигур. Каждая фигура представленным объектом **[фигуры](shape-object-publisher.md)** . С помощью метода **[Item](groupshapes-item-method-publisher.md)** с объектом, можно работать с одним фигур в группе без необходимости их разгруппировать.
 


## <a name="example"></a>Пример

Свойство **[GroupItems](shape-groupitems-property-publisher.md)** используется для возврата коллекции **GroupShapes** . Используйте **GroupItems** (индекс), где индекса — это число отдельные фигуры в группы фигур, чтобы получить одну из коллекции **GroupShapes** . Следующий пример добавляет три треугольники в активный документ, группирует их, задает цвет для всей группы и затем меняет свой цвет для третьего треугольник только.
 

 

```
Sub WorkWithGroupShapes() 
 With ActiveDocument.Pages.Add(Count:=1, After:=1).Shapes 
 .AddShape(msoShapeIsoscelesTriangle, _ 
 50, 50, 100, 100).Name = "shpOne" 
 .AddShape(msoShapeIsoscelesTriangle, _ 
 200, 50, 100, 100).Name = "shpTwo" 
 .AddShape(msoShapeIsoscelesTriangle, _ 
 350, 50, 100, 100).Name = "shpThree" 
 With .Range(Array("shpOne", "shpTwo", "shpThree")).Group 
 .Fill.PresetTextured PresetTexture:=msoTextureBlueTissuePaper 
 .GroupItems(3).Fill.PresetTextured _ 
 PresetTexture:=msoTextureGreenMarble 
 End With 
 End With 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Элемент](groupshapes-item-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](groupshapes-application-property-publisher.md)|
|[Count](groupshapes-count-property-publisher.md)|
|[Родительский раздел](groupshapes-parent-property-publisher.md)|

