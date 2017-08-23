---
title: "Свойство Table.Rows (издатель)"
keywords: vbapb10.chm4784134
f1_keywords: vbapb10.chm4784134
ms.prod: publisher
api_name: Publisher.Table.Rows
ms.assetid: 97a543b9-a1d7-c7f8-9f3c-e08256e0b364
ms.date: 06/08/2017
ms.openlocfilehash: 9b760c695b85a3720c29d91827052658f13004a2
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="tablerows-property-publisher"></a>Свойство Table.Rows (издатель)

Возвращает набор **[строк](rows-object-publisher.md)** , который представляет все строки в таблице в диапазон, выделенный фрагмент или таблицу.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Строк**

 переменная _expression_A, представляет собой объект- **таблицы** .


## <a name="remarks"></a>Заметки

Сведения о возврате один элемент коллекции видеть [возврата объекта из коллекции](returning-an-object-from-a-collection-publisher.md).


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


