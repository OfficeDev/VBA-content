---
title: "Метод TextRange.Select (издатель)"
keywords: vbapb10.chm5308457
f1_keywords: vbapb10.chm5308457
ms.prod: publisher
api_name: Publisher.TextRange.Select
ms.assetid: 36097502-2b06-37ac-3148-43a82cca4411
ms.date: 06/08/2017
ms.openlocfilehash: e26b9717c1a0e6a77c5d4723955d1ecc74290540
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangeselect-method-publisher"></a>Метод TextRange.Select (издатель)

Выбирает указанный объект.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Выберите**

 переменная _expression_A, представляющий объект **TextRange** .


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


