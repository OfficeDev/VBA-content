---
title: "Метод CellRange.Select (издатель)"
keywords: vbapb10.chm5177353
f1_keywords: vbapb10.chm5177353
ms.prod: publisher
api_name: Publisher.CellRange.Select
ms.assetid: 15b0fc0b-8cac-9ff9-bac3-cf15351c7645
ms.date: 06/08/2017
ms.openlocfilehash: 72c6b12a368560c0b4669573d930d04524d491b5
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="cellrangeselect-method-publisher"></a>Метод CellRange.Select (издатель)

Выбирает указанный объект.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Выберите**

 переменная _expression_A, представляет собой объект- **CellRange** .


## <a name="example"></a>Пример

В этом примере выбирает левый верхний угол из таблицы, который был добавлен к первой страницы в активной публикации.


```vb
Dim shpTable As Shape 
Dim cllTemp As Cell 
 
With ActiveDocument.Pages(1).Shapes 
 Set shpTable = .AddTable(NumRows:=3, NumColumns:=3, _ 
 Left:=100, Top:=100, Width:=150, Height:=150) 
 Set cllTemp = shpTable.Table.Cells.Item(1) 
 cllTemp.Select 
End With
```

В этом примере выбирает первый столбец из таблицы, который был добавлен к первой страницы в активной публикации.




```vb
Dim shpTable As Shape 
Dim crTemp As CellRange 
 
With ActiveDocument.Pages(1).Shapes 
 Set shpTable = .AddTable(NumRows:=3, NumColumns:=3, _ 
 Left:=100, Top:=100, Width:=150, Height:=150) 
 Set crTemp = shpTable.Table.Cells(StartRow:=1, _ 
 StartColumn:=1, EndRow:=3, EndColumn:=1) 
 crTemp.Select 
End With
```

В этом примере выбирает первые пять знаков в форму одно на странице один из активных публикации.




```vb
ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.Characters(1, 5).Select
```


