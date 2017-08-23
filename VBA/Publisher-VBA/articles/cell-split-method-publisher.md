---
title: "Метод Cell.Split (издатель)"
keywords: vbapb10.chm5111844
f1_keywords: vbapb10.chm5111844
ms.prod: publisher
api_name: Publisher.Cell.Split
ms.assetid: 99bc34df-c8dc-90e5-4262-dbe0a9c9b61d
ms.date: 06/08/2017
ms.openlocfilehash: 8f50f42a794b9e2da2331ba308bd42fd24bdfbdf
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="cellsplit-method-publisher"></a>Метод Cell.Split (издатель)

Разделение объединенных ячеек обратно в его составные ячеек. Возвращает объект **[CellRange](cellrange-object-publisher.md)** , представляющий целостные ячеек.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Разделение**

 переменная _expression_A, представляет собой объект- **ячейки** .


### <a name="return-value"></a>Возвращаемое значение

CellRange


## <a name="remarks"></a>Заметки

Если указанный ячейки не объединенные ячейки из с помощью метода **[объединения](cell-merge-method-publisher.md)** , возникает ошибка.


## <a name="example"></a>Пример

Следующий пример разделяет первую ячейку в таблице в один фигуры на странице один из активных публикации в его составные ячейки. Фигура один должен содержать таблицы, первой ячейки из которых является объединенные ячейки для работы этого примера.


```vb
Dim cllMerged As Cell 
 
Set cllMerged = ActiveDocument.Pages(1).Shapes(1).Table.Cells.Item(1) 
 
cllMerged.Split
```


