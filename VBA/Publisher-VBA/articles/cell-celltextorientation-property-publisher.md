---
title: "Свойство Cell.CellTextOrientation (издатель)"
keywords: vbapb10.chm5111845
f1_keywords: vbapb10.chm5111845
ms.prod: publisher
api_name: Publisher.Cell.CellTextOrientation
ms.assetid: ad2c2f15-358c-7bbc-b249-b886a99ea4a5
ms.date: 06/08/2017
ms.openlocfilehash: 93bc9ca3ff0001f662776b5529abb6394e116ab5
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="cellcelltextorientation-property-publisher"></a>Свойство Cell.CellTextOrientation (издатель)

Возвращает или задает **PbTextOrientation** , представляющий поток текст в указанной ячейке таблицы. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **CellTextOrientation**

 переменная _expression_A, представляет собой объект- **ячейки** .


### <a name="return-value"></a>Возвращаемое значение

PbTextOrientation


## <a name="remarks"></a>Заметки

Значение свойства **CellTextOrientation** может иметь одно из **[PbTextOrientation](pbtextorientation-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

В этом примере увеличивается высота ячеек в первой строке и затем добавляет текст по вертикали ориентированного заголовка.


```vb
Sub VerticalText() 
 Dim rowTable As Row 
 Dim celTable As Cell 
 
 With ActiveDocument.Pages(2).Shapes(1).Table.Rows(1) 
 .Height = Application.InchesToPoints(1.5) 
 For Each celTable In .Cells 
 With celTable 
 .CellTextOrientation _ 
 = pbTextOrientationVerticalEastAsia 
 .TextRange.ParagraphFormat.Alignment _ 
 = pbParagraphAlignmentCenter 
 .TextRange.Text = "Column Heading " _ 
 &; celTable.Column 
 End With 
 Next 
 End With 
End Sub
```


