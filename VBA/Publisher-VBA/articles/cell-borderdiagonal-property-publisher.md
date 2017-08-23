---
title: "Свойство Cell.BorderDiagonal (издатель)"
keywords: vbapb10.chm5111810
f1_keywords: vbapb10.chm5111810
ms.prod: publisher
api_name: Publisher.Cell.BorderDiagonal
ms.assetid: 2c857a1b-2a0f-5796-9397-ad113dd984cb
ms.date: 06/08/2017
ms.openlocfilehash: 2b485afaa70badad82cbd63c95ca8cf9cdda58da
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="cellborderdiagonal-property-publisher"></a>Свойство Cell.BorderDiagonal (издатель)

Возвращает объект [CellBorder](cellborder-object-publisher.md), представляющий косую границу для указанной ячейке таблицы.


## <a name="syntax"></a>Синтаксис

 _выражение_. **BorderDiagonal**

 переменная _expression_A, представляет собой объект- **ячейки** .


### <a name="return-value"></a>Возвращаемое значение

CellBorder


## <a name="example"></a>Пример

В этом примере диагонали разделяет каждой ячейки в указанную таблицу и добавляет косую границу. В этом примере предполагается, что первую фигуру на вторую страницу — это таблица и не другого типа фигуры.


```vb
Sub FillCellsByRow() 
 Dim shpTable As Shape 
 Dim rowTable As Row 
 Dim celTable As Cell 
 Dim intCell As Integer 
 
 intCell = 1 
 
 Set shpTable = ActiveDocument.Pages(2).Shapes(1) 
 For Each rowTable In shpTable.Table.Rows 
 For Each celTable In rowTable.Cells 
 If intCell Mod 2 = 0 Then 
 With celTable 
 .Diagonal = pbTableCellDiagonalDown 
 With .BorderDiagonal 
 .Weight = 1 
 .Color.RGB = RGB(Red:=0, Green:=0, Blue:=0) 
 End With 
 End With 
 celTable.Fill.ForeColor.RGB = RGB _ 
 (Red:=180, Green:=180, Blue:=180) 
 Else 
 celTable.Fill.ForeColor.RGB = RGB _ 
 (Red:=255, Green:=255, Blue:=255) 
 End If 
 intCell = intCell + 1 
 Next celTable 
 Next rowTable 
 
End Sub
```


