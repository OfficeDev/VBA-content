---
title: "Свойство Selection.ShapeRange (издатель)"
keywords: vbapb10.chm851972
f1_keywords: vbapb10.chm851972
ms.prod: publisher
api_name: Publisher.Selection.ShapeRange
ms.assetid: d95cce6d-e3a2-09b9-a6d5-749e0476544c
ms.date: 06/08/2017
ms.openlocfilehash: a875df0592b659f8ceae4d980954e04130fd6630
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="selectionshaperange-property-publisher"></a>Свойство Selection.ShapeRange (издатель)

Возвращает коллекцию **[ShapeRange](shaperange-object-publisher.md)** , представляющий все объекты **фигуры** в указанный диапазон или выделить фрагмент. Диапазон фигура может содержать рисунки, фигуры, рисунки, OLE объекты, элементы управления ActiveX, текстовые объекты и выноски.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ShapeRange**

 переменная _expression_A, представляющий объект **Selection** .


### <a name="return-value"></a>Возвращаемое значение

ShapeRange


## <a name="example"></a>Пример

В следующем примере задается узор заливки для всех фигур в выделение. В этом примере предполагается, что один или несколько фигур выбраны в активной публикации.


```vb
Sub ChangeFillForShapeRange() 
 Selection.ShapeRange.Fill.Patterned Pattern:=msoPattern20Percent 
End Sub
```

В следующем примере применяется тени и заливки форматирования для всех фигур в выделение. В этом примере предполагается, что один или несколько фигур выбраны в активной публикации.




```vb
Sub SetShadowForSelectedShapes() 
 With Selection.ShapeRange 
 .Shadow.Type = msoShadow6 
 .Fill.Patterned Pattern:=msoPatternDottedDiamond 
 End With 
End Sub
```


