---
title: "Свойство CellRange.Column (издатель)"
keywords: vbapb10.chm5177346
f1_keywords: vbapb10.chm5177346
ms.prod: publisher
api_name: Publisher.CellRange.Column
ms.assetid: 77925e68-c8ff-9732-32c4-4f082eb3fd1c
ms.date: 06/08/2017
ms.openlocfilehash: abf461f97eaa135164e0af6fa48ffc4a80f40923
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="cellrangecolumn-property-publisher"></a>Свойство CellRange.Column (издатель)

Возвращает значение типа **Long** , который представляет столбец таблицы с указанной ячейке. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Столбец**

 переменная _expression_A, представляет собой объект- **CellRange** .


## <a name="example"></a>Пример

В этом примере добавляет страницу в активной публикации, создается таблица на новой странице и диагонали разделяет всем ячейкам в четных столбцов.


```vb
Sub CreateNewTable() 
 
 Dim pgeNew As Page 
 Dim shpTable As Shape 
 Dim tblNew As Table 
 Dim celTable As Cell 
 Dim rowTable As Row 
 
 'Creates a new document with a five-row by five-column table 
 Set pgeNew = ActiveDocument.Pages.Add(Count:=1, After:=1) 
 Set shpTable = pgeNew.Shapes.AddTable(NumRows:=5, NumColumns:=5, _ 
 Left:=72, Top:=72, Width:=468, Height:=100) 
 Set tblNew = shpTable.Table 
 
 'Inserts a diagonal split into all cells in even-numbered columns 
 For Each rowTable In tblNew.Rows 
 For Each celTable In rowTable.Cells 
 If celTable.Column Mod 2 = 0 Then 
 celTable.Diagonal = pbTableCellDiagonalUp 
 End If 
 Next celTable 
 Next rowTable 
 
End Sub
```


