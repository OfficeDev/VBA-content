---
title: "Свойство Row.Cells (издатель)"
keywords: vbapb10.chm4849666
f1_keywords: vbapb10.chm4849666
ms.prod: publisher
api_name: Publisher.Row.Cells
ms.assetid: 2a866890-d564-b9bc-c553-06669f376788
ms.date: 06/08/2017
ms.openlocfilehash: 3c4c428783be62b0d9704c81d13b4cf8ffdf2517
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="rowcells-property-publisher"></a>Свойство Row.Cells (издатель)

Возвращает объект **[CellRange](cellrange-object-publisher.md)** , представляющий одну или несколько ячеек в строке таблицы.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Ячейки**

 переменная _expression_A, представляет собой объект- **строку** .


## <a name="example"></a>Пример

В этом примере выполняется объединение ячеек первой и второй в первый столбец указанную таблицу.


```vb
Sub MergeCell() 
 With ActiveDocument.Pages(1).Shapes(2).Table.Columns(1) 
 .Cells(1).Merge MergeTo:=.Cells(2) 
 End With 
End Sub
```

В этом примере применяется структуры толстой границей в первую ячейку в столбце второй указанную таблицу.




```vb
Sub OutlineBorderCell() 
 With ActiveDocument.Pages(1).Shapes(2).Table.Columns(2).Cells(1) 
 .BorderLeft.Weight = 5 
 .BorderRight.Weight = 5 
 .BorderTop.Weight = 5 
 .BorderBottom.Weight = 5 
 End With 
End Sub
```


