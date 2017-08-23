---
title: "Свойство Cell.Diagonal (издатель)"
keywords: vbapb10.chm5111816
f1_keywords: vbapb10.chm5111816
ms.prod: publisher
api_name: Publisher.Cell.Diagonal
ms.assetid: 4ec93690-38ef-7434-55a5-419f14c9ea73
ms.date: 06/08/2017
ms.openlocfilehash: d9945c4562db5e4474cb6f2910a34c5263f9c117
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="celldiagonal-property-publisher"></a>Свойство Cell.Diagonal (издатель)

Задает или возвращает константу **PbCellDiagonalType** , представляющий ячейку, в которой диагонали разделение. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Диагональный**

 переменная _expression_A, представляет собой объект- **ячейки** .


### <a name="return-value"></a>Возвращаемое значение

PbCellDiagonalType


## <a name="remarks"></a>Заметки

**Диагональные** значение свойства может иметь одно из **[PbCellDiagonalType](pbcelldiagonaltype-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.


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


