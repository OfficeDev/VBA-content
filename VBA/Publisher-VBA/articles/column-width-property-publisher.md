---
title: "Свойство Column.Width (издатель)"
keywords: vbapb10.chm4980739
f1_keywords: vbapb10.chm4980739
ms.prod: publisher
api_name: Publisher.Column.Width
ms.assetid: 9596b828-a5ce-e501-db59-a0e1533108b3
ms.date: 06/08/2017
ms.openlocfilehash: cd945078ee3d1e8cf142a0da45f771401aa07bfc
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="columnwidth-property-publisher"></a>Свойство Column.Width (издатель)

Возвращает или задает **Variant** , который представляет ширину (в точках) указанного столбца или фигуры. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Ширина**

 переменная _expression_A, представляет собой объект- **столбец** .


## <a name="example"></a>Пример

В этом примере создается новая таблица и задает высоту и ширину второй строк и столбцов, соответственно.


```vb
Sub SetRowHeightColumnWidth() 
 With ActiveDocument.Pages(1).Shapes.AddTable(NumRows:=3, _ 
 NumColumns:=3, Left:=80, Top:=80, Width:=400, Height:=12).Table 
 .Rows(2).Height = 72 
 .Columns(2).Width = 72 
 End With 
End Sub
```


