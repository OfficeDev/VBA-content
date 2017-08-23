---
title: "Свойство CellRange.Row (издатель)"
keywords: vbapb10.chm5177350
f1_keywords: vbapb10.chm5177350
ms.prod: publisher
api_name: Publisher.CellRange.Row
ms.assetid: ac5bccf0-6c9b-ce0e-20e5-f06ef29886c6
ms.date: 06/08/2017
ms.openlocfilehash: 3bc1a6e683e2e30db9983710983113ed0bb53580
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="cellrangerow-property-publisher"></a>Свойство CellRange.Row (издатель)

Возвращает значение типа **Long** , представляющее номер строки, содержащей указанную ячейку. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Строка**

 переменная _expression_A, представляет собой объект- **CellRange** .


## <a name="example"></a>Пример

В этом примере вводит заливки для всех четных строк и очищает заливки для всех нечетных строк в указанной таблице. В этом примере предполагается, что указанные форму — это таблица и не другого типа фигуры.


```vb
Sub FillCellsByRow() 
 Dim shpTable As Shape 
 Dim rowTable As Row 
 Dim celTable As Cell 
 
 Set shpTable = ActiveDocument.Pages(1).Shapes _ 
 .AddTable(NumRows:=5, NumColumns:=5, Left:=100, _ 
 Top:=100, Width:=100, Height:=12) 
 For Each rowTable In shpTable.Table.Rows 
 For Each celTable In rowTable.Cells 
 If celTable.Row Mod 2 = 0 Then 
 celTable.Fill.ForeColor.RGB = RGB _ 
 (Red:=180, Green:=180, Blue:=180) 
 Else 
 celTable.Fill.ForeColor.RGB = RGB _ 
 (Red:=255, Green:=255, Blue:=255) 
 End If 
 Next celTable 
 Next rowTable 
 
End Sub
```


