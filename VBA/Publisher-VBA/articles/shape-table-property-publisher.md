---
title: "Свойство Shape.Table (издатель)"
keywords: vbapb10.chm2228328
f1_keywords: vbapb10.chm2228328
ms.prod: publisher
api_name: Publisher.Shape.Table
ms.assetid: a9b29d1f-2459-556c-56f8-f8f809b879c9
ms.date: 06/08/2017
ms.openlocfilehash: 7f40a4190a50108820f53ceff5e99487cd518ee9
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapetable-property-publisher"></a>Свойство Shape.Table (издатель)

Возвращает объект **таблицы** , который представляет таблицу в Microsoft Publisher.


## <a name="syntax"></a>Синтаксис

 _выражение_. **В таблице**

 переменная _expression_A, представляющий объект **фигуры** .


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


