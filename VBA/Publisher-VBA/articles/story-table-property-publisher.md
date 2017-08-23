---
title: "Свойство Story.Table (издатель)"
keywords: vbapb10.chm5832710
f1_keywords: vbapb10.chm5832710
ms.prod: publisher
api_name: Publisher.Story.Table
ms.assetid: e9da80d3-ea3c-b47c-d434-498c72955c14
ms.date: 06/08/2017
ms.openlocfilehash: 6f39ff1ee6d3d010ecbbfe1ff148109272cf8307
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="storytable-property-publisher"></a>Свойство Story.Table (издатель)

Возвращает объект **таблицы** , который представляет таблицу в Microsoft Publisher.


## <a name="syntax"></a>Синтаксис

 _выражение_. **В таблице**

 переменная _expression_A, представляет собой объект- **материала** .


## <a name="example"></a>Пример

В следующем примере добавляется таблица 5 x 5 на первой странице active публикации и затем выбирает первый столбец новой таблицы.


```vb
Sub NewTable() 
 With ActiveDocument.Pages(1).Shapes.AddTable(NumRows:=5, _ 
 NumColumns:=5, Left:=72, Top:=300, Width:=400, Height:=100) 
 .Table.Columns(3).Cells(3).Fill.ForeColor.RGB = RGB _ 
 (Red:=255, Green:=0, Blue:=0) 
 End With 
End Sub
```

В следующем примере выбирается указанную таблицу в активной публикации. В этом примере предполагает наличие по крайней мере один фигуры на первой странице active публикации.




```vb
Sub SelectTable() 
 With ActiveDocument.Pages(1).Shapes(1) 
 If .Type = pbTable Then 
 .Table.Rows(3).Cells(3).Fill.ForeColor _ 
 .RGB = RGB(Red:=150, Green:=150, Blue:=150) 
 End If 
 End With 
End Sub
```


