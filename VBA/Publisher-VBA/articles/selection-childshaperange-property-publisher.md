---
title: "Свойство Selection.ChildShapeRange (издатель)"
keywords: vbapb10.chm851973
f1_keywords: vbapb10.chm851973
ms.prod: publisher
api_name: Publisher.Selection.ChildShapeRange
ms.assetid: 8ef96e85-2f25-7b3a-4465-7e22fdbbaa9a
ms.date: 06/08/2017
ms.openlocfilehash: 8a3d36f697414cc979d04e8130ecdbdcc344fce4
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="selectionchildshaperange-property-publisher"></a>Свойство Selection.ChildShapeRange (издатель)

Возвращает объект **[ShapeRange](shaperange-object-publisher.md)** , представляющий фигур дочерние выделения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ChildShapeRange**

 переменная _expression_A, представляющий объект **Selection** .


### <a name="return-value"></a>Возвращаемое значение

ShapeRange


## <a name="example"></a>Пример

В этом примере создается новая страница в активной публикации, заполняет страница с фигурами и выбирает и группы фигур. Затем после отмены выбора двух фигур группы, он изменяется типа автофигуры для одну из фигур.


```vb
Sub ChangeFillToChildShape() 
 
 With ThisDocument.Pages(1) 
 With .Shapes 
 .AddShape msoShape4pointStar, 10, 10, 175, 175 
 .AddShape msoShapeOval, 100, 100, 175, 75 
 .AddShape msoShapeOval, 150, 150, 175, 75 
 .Range.Group 
 .SelectAll 
 End With 
 .Shapes(1).GroupItems(1).Select msoFalse 
 .Shapes(1).GroupItems(2).Select msoFalse 
 End With 
 
 Selection.ChildShapeRange(3).AutoShapeType = msoShapeDiamond 
 
End Sub
```


