---
title: "Метод Cell.Merge (издатель)"
keywords: vbapb10.chm5111842
f1_keywords: vbapb10.chm5111842
ms.prod: publisher
api_name: Publisher.Cell.Merge
ms.assetid: 09ae6910-ba47-4b0e-5792-ac9eb1298d57
ms.date: 06/08/2017
ms.openlocfilehash: 91841f8342bf763bdb4ddab402d9645951e28394
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="cellmerge-method-publisher"></a>Метод Cell.Merge (издатель)

Объединяет указанной ячейке таблицы с другой ячейки. Результатом является одну ячейку.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Объединение** ( **_MergeTo_**)

 переменная _expression_A, представляет собой объект- **ячейки** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|MergeTo|Обязательное свойство.| **Ячейки**|Ячейки для объединения. должен быть рядом с указанной ячейке или ошибка возникает.|

## <a name="example"></a>Пример

В этом примере выполняется объединение первых двух смежных ячеек первого столбца указанную таблицу.


```vb
Sub MergeCell() 
 With ActiveDocument.Pages(1).Shapes(2).Table 
 .Rows(1).Cells(1).Merge MergeTo:=.Rows(2).Cells(1) 
 End With 
End Sub
```


