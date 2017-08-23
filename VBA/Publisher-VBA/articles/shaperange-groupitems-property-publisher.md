---
title: "Свойство ShapeRange.GroupItems (издатель)"
keywords: vbapb10.chm2293816
f1_keywords: vbapb10.chm2293816
ms.prod: publisher
api_name: Publisher.ShapeRange.GroupItems
ms.assetid: d37c75cd-a651-51d1-42c7-59879ccbbf1d
ms.date: 06/08/2017
ms.openlocfilehash: 1821fcb9ebc2085504f0a90f83cd9022a4855459
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangegroupitems-property-publisher"></a>Свойство ShapeRange.GroupItems (издатель)

Если указанные форму — это группа, возвращает коллекцию **[GroupShapes](groupshapes-object-publisher.md)** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **GroupItems**

 переменная _expression_A, представляющий объект **ShapeRange** .


## <a name="remarks"></a>Заметки

Все объекты смарт-будет рассматриваться как группы фигур.


## <a name="example"></a>Пример

В этом примере добавляется три треугольники на публикацию, группирует их, задает цвет для всей группы и затем меняет свой цвет для второй треугольник только.


```vb
Sub Grouper() 
 
 Dim docSheet As Document 
 
 Set docSheet = ActiveDocument 
 With docSheet.MasterPages.Item(1).Shapes 
 ' Add the 3 triangles 
 .AddShape(Type:=msoShapeIsoscelesTriangle, _ 
 Left:=10, Top:=10, Width:=100, Height:=100).Name = "shpOne" 
 .AddShape(Type:=msoShapeIsoscelesTriangle, _ 
 Left:=150, Top:=10, Width:=100, Height:=100).Name = "shpTwo" 
 .AddShape(Type:=msoShapeIsoscelesTriangle, _ 
 Left:=300, Top:=10, Width:=100, Height:=100).Name = "shpThree" 
 ' Group and fill the 3 triangles 
 With .Range(Array("shpOne", "shpTwo", "shpThree")).Group 
 .Fill.PresetTextured msoTextureBlueTissuePaper 
 .GroupItems(2).Fill.PresetTextured msoTextureGreenMarble 
 End With 
 End With 
 
End Sub
```


