---
title: "Свойство Table.Columns (издатель)"
keywords: vbapb10.chm4784131
f1_keywords: vbapb10.chm4784131
ms.prod: publisher
api_name: Publisher.Table.Columns
ms.assetid: fb55ba62-64a4-2221-3cc7-b349dc2f6934
ms.date: 06/08/2017
ms.openlocfilehash: 616af2bc18298fe9317529b4857e54a5b3e87e77
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="tablecolumns-property-publisher"></a>Свойство Table.Columns (издатель)

Возвращает коллекцию **[столбцов](columns-object-publisher.md)** , представляющий все столбцы указанной таблицы.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Столбцы**

 переменная _expression_A, представляет собой объект- **таблицы** .


## <a name="example"></a>Пример

В этом примере вводит полужирный номер в каждой ячейки в указанной таблице. В этом примере предполагается, что указанные форму — это таблица и не другого типа фигуры.


```vb
Sub CountCellsByColumn() 
 Dim shpTable As Shape 
 Dim colTable As Column 
 Dim celTable As Cell 
 Dim intCount As Integer 
 
 intCount = 1 
 
 Set shpTable = ActiveDocument.Pages(2).Shapes(1) 
 
 'Loops through each column in the table 
 For Each colTable In shpTable.Table.Columns 
 
 'Loops through each cell in the column 
 For Each celTable In colTable.Cells 
 With celTable.Text 
 .Text = intCount 
 .ParagraphFormat.Alignment = _ 
 pbParagraphAlignmentCenter 
 .Font.Bold = msoTrue 
 intCount = intCount + 1 
 End With 
 Next celTable 
 Next colTable 
 
End Sub
```


