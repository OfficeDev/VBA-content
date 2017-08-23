---
title: "Свойство Cell.BorderLeft (издатель)"
keywords: vbapb10.chm5111812
f1_keywords: vbapb10.chm5111812
ms.prod: publisher
api_name: Publisher.Cell.BorderLeft
ms.assetid: f996a96f-4392-48c2-e5c2-bfe373a7997a
ms.date: 06/08/2017
ms.openlocfilehash: 9a306c640ad1e3e163b830ebdb52f0df0a045f08
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="cellborderleft-property-publisher"></a>Свойство Cell.BorderLeft (издатель)

Возвращает объект [CellBorder](cellborder-object-publisher.md), представляющий левую границу для указанной ячейке таблицы.


## <a name="syntax"></a>Синтаксис

 _выражение_. **BorderLeft**

 переменная _expression_A, представляет собой объект- **ячейки** .


### <a name="return-value"></a>Возвращаемое значение

CellBorder


## <a name="example"></a>Пример

В этом примере создается шашками, теперь разработки, с помощью границы и цвет заливки с помощью существующей таблицы. Предполагается первую фигуру на вторую страницу таблицы и не другого типа фигуры и таблицы на наличие нечетного числа столбцов.


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
 With celTable 
 With .BorderBottom 
 .Weight = 2 
 .Color.RGB = RGB(Red:=0, Green:=0, Blue:=0) 
 End With 
 With .BorderTop 
 .Weight = 2 
 .Color.RGB = RGB(Red:=0, Green:=0, Blue:=0) 
 End With 
 With .BorderLeft 
 .Weight = 2 
 .Color.RGB = RGB(Red:=0, Green:=0, Blue:=0) 
 End With 
 With .BorderRight 
 .Weight = 2 
 .Color.RGB = RGB(Red:=0, Green:=0, Blue:=0) 
 End With 
 End With 
 If intCell Mod 2 = 0 Then 
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


