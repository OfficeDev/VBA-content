---
title: "Метод CatalogMergeShapes.Range (издатель)"
keywords: vbapb10.chm8388612
f1_keywords: vbapb10.chm8388612
ms.prod: publisher
api_name: Publisher.CatalogMergeShapes.Range
ms.assetid: e92dcac4-4694-8a22-61da-09fcd98c72ce
ms.date: 06/08/2017
ms.openlocfilehash: 7a53027b6101bb4b2d1427110e37a52cfd0de49d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="catalogmergeshapesrange-method-publisher"></a>Метод CatalogMergeShapes.Range (издатель)

Возвращает объект **[ShapeRange](shaperange-object-publisher.md)** , который представляет собой подмножество фигуры в коллекции **фигур** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **Диапазон** ( **_Индекс_**)

 переменная _expression_A, представляет собой объект- **CatalogMergeShapes** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Индекс|Обязательное свойство.| **Variant**|Отдельные фигуры, которые должны быть включены в диапазоне. Может быть целое число, указывающее индекс фигуры, string, указывающее имя фигуры или массив, содержащий целых значений или строк. Если индекс указан, метод **диапазона** возвращает все объекты в указанном семействе сайтов.|

### <a name="return-value"></a>Возвращаемое значение

ShapeRange


## <a name="example"></a>Пример

Чтобы указать массив целых значений или строк для **_индекса_**, можно использовать функцию **массива** . Например следующие инструкции возвращает двумя фигурами, указанный в параметре name.


```vb
Dim arrShapes As Variant 
Dim shpRange As ShapeRange 
 
Set arrShapes = Array("Oval 4", "Rectangle 5") 
Set shpRange = ActiveDocument.Pages(1) _ 
 .Shapes.Range(arrShapes)
```

В этом примере задается узор заливки для фигур одним и три active публикацией.




```vb
ActiveDocument.Pages(1).Shapes.Range(Array(1, 3)).Fill _ 
 .Patterned msoPatternHorizontalBrick
```




```

```

В этом примере задается узор заливки для фигуры, с именем «Овал 4» и «прямоугольник 5" на первой странице.




```vb
Dim arrShapes As Variant 
Dim shpRange As ShapeRange 
 
arrShapes = Array("Oval 4", "Rectangle 5") 
 
Set shpRange = ActiveDocument.Pages(1).Shapes.Range(arrShapes) 
 
shpRange.Fill.Patterned msoPatternHorizontalBrick
```

В этом примере задается узор заливки для всех фигур на первой странице.




```vb
ActiveDocument.Pages(1).Shapes _ 
 .Range.Fill.Patterned msoPatternHorizontalBrick
```

В этом примере задается узор заливки для фигуры одно на первой странице.




```vb
Dim shpRange As ShapeRange 
 
Set shpRange = ActiveDocument.Pages(1).Shapes.Range(1) 
 
shpRange.Fill.Patterned msoPatternHorizontalBrick
```

В этом примере создается массив, содержащий все автофигуры на первой странице, этот массив используется для определения диапазона фигуры и затем распределяет всех фигур в этот диапазон по горизонтали.




```vb
Dim numShapes As Long 
Dim numAutoShapes As Long 
Dim autoShpArray As Variant 
Dim intLoop As Integer 
Dim shpRange As ShapeRange 
 
With ActiveDocument.Pages(1).Shapes 
 
 numShapes = .Count 
 If numShapes > 1 Then 
 
 numAutoShapes = 0 
 ReDim autoShpArray(1 To numShapes) 
 
 For intLoop = 1 To numShapes 
 If .Item(intLoop).Type = msoAutoShape Then 
 numAutoShapes = numAutoShapes + 1 
 autoShpArray(numAutoShapes) = .Item(intLoop).Name 
 End If 
 Next 
 
 If numAutoShapes > 1 Then 
 ReDim Preserve autoShpArray(1 To numAutoShapes) 
 Set shpRange = .Range(autoShpArray) 
 shpRange.Distribute _ 
 DistributeCmd:=msoDistributeHorizontally, _ 
 RelativeTo:=False 
 End If 
 
 End If 
 
End With
```


