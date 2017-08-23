---
title: "Метод CellRange.Merge (издатель)"
keywords: vbapb10.chm5177352
f1_keywords: vbapb10.chm5177352
ms.prod: publisher
api_name: Publisher.CellRange.Merge
ms.assetid: f097659c-d1b8-f2bb-c4fc-5efc2b7417dd
ms.date: 06/08/2017
ms.openlocfilehash: 83b0ddafb081e2f18b55fb432a4a1898c77e9f3f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="cellrangemerge-method-publisher"></a>Метод CellRange.Merge (издатель)

Объединение ячеек указанной таблицы друг с другом. Результатом является одну ячейку.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Объединение**

 переменная _expression_A, представляет собой объект- **CellRange** .


## <a name="example"></a>Пример

В этом примере выполняется объединение первых двух смежных ячеек в первых двух строк в указанной таблице.


```vb
Sub MergeCells() 
 ActiveDocument.Pages(1).Shapes(2).Table _ 
 .Cells(StartRow:=1, StartColumn:=1, _ 
 EndRow:=2, EndColumn:=2).Merge 
End Sub
```


