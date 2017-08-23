---
title: "Свойство ShapeRange.Vertices (издатель)"
keywords: vbapb10.chm2293845
f1_keywords: vbapb10.chm2293845
ms.prod: publisher
api_name: Publisher.ShapeRange.Vertices
ms.assetid: 0beb2323-8db6-c8c2-2f34-4c1ffde7fddc
ms.date: 06/08/2017
ms.openlocfilehash: 2dfccff7ff7deae6ae738f4826a70684222eb877
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangevertices-property-publisher"></a>Свойство ShapeRange.Vertices (издатель)

Возвращает координаты вершин указанного freeform документа (и контрольные точки для кривых Безье) в формате пары координат. Только для чтения **Variant**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Вершины**

 переменная _expression_A, представляющий объект **ShapeRange** .


## <a name="remarks"></a>Заметки

Можно использовать массива, возвращаемого этим свойством как аргумент для методов [AddCurve](shapes-addcurve-method-publisher.md)или [AddPolyline](shapes-addpolyline-method-publisher.md).

В следующей таблице показано, как свойство **вершин** связывает значения в массиве `vertArray()` с координаты вершины треугольника.



|**элемент vertArray**|**Contains**|
|:-----|:-----|
| `vertArray(1, 1)`|Расстояние по горизонтали из первой вершины в левой части страницы.|
| `vertArray(1, 2)`|Расстояние по вертикали из первой вершины в верхней части страницы.|
| `vertArray(2, 1)`|Расстояние по горизонтали от второй вершины в левой части страницы.|
| `vertArray(2, 2)`|Расстояние по вертикали от второй вершины в верхней части страницы.|
| `vertArray(3, 1)`|Расстояние по горизонтали от третьей вершины в левой части страницы.|
| `vertArray(3, 2)`|Расстояние по вертикали от третьей вершины в верхней части страницы.|

## <a name="example"></a>Пример

В этом примере присваивает переменной массива координаты вершин фигуры один активный публикации `vertArray()` и отображает координаты для первой вершины.


```vb
Dim vertArray As Variant 
Dim sngX1 As Single 
Dim sngY1 As Single 
 
With ActiveDocument.Pages(1).Shapes(1) 
 vertArray = .Vertices 
 sngX1 = vertArray(1, 1) 
 sngY1 = vertArray(1, 2) 
 MsgBox "First vertex coordinates: " &; sngX1 &; ", " &; sngY1 
End With
```

В этом примере создается график, который выполняет ту же функцию геометрические как один фигуры в активной публикации. Фигура один должен содержать 3n + 1 вершины для этого примера, где n — целое число, большее или равное 1.




```vb
With ActiveDocument.Pages(1).Shapes 
 .AddCurve SafeArrayOfPoints:=.Item(1).Vertices 
End With 

```


