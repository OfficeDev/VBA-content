---
title: "Свойство ShapeNode.Points (издатель)"
keywords: vbapb10.chm3539201
f1_keywords: vbapb10.chm3539201
ms.prod: publisher
api_name: Publisher.ShapeNode.Points
ms.assetid: 30235d5a-9f05-4cc4-f62f-ac3cf4916e0d
ms.date: 06/08/2017
ms.openlocfilehash: 2d58ed4841a78cdc8467708fd15137768cf04268
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapenodepoints-property-publisher"></a>Свойство ShapeNode.Points (издатель)

Получает координаты _x_ и _y_ узла фигуры. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Точки**

 переменная _expression_A, представляет собой объект- **ShapeNode** .


## <a name="remarks"></a>Заметки

Это свойство доступно только для чтения. Используйте метод **[SetPosition](shapenodes-setposition-method-publisher.md)** , чтобы задать расположение узла.


## <a name="example"></a>Пример

В этом примере перемещает два узла в один фигуры на первой странице активная публикация 200 точек вправо и вниз 300 точек. В данном примере для работы фигуры один должен быть freeform документа.


```vb
Sub SetPointsPosition() 
 Dim varArray As Variant 
 Dim intX As Integer 
 Dim intY As Integer 
 With ActiveDocument.Pages(1).Shapes(1).Nodes 
 varArray = .Item(2).Points 
 intX = varArray(1, 1) 
 intY = varArray(1, 2) 
 .SetPosition Index:=2, X1:=intX + 200, Y1:=intY + 300 
 End With 
End Sub
```


