---
title: "Свойство Cell.HasText (издатель)"
keywords: vbapb10.chm5111824
f1_keywords: vbapb10.chm5111824
ms.prod: publisher
api_name: Publisher.Cell.HasText
ms.assetid: b44c5d24-7ac1-a63d-6986-05ed9c91dd8e
ms.date: 06/08/2017
ms.openlocfilehash: 2992bd4e2989bd232f5a6cc4ee1500602f97b5f9
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="cellhastext-property-publisher"></a>Свойство Cell.HasText (издатель)

Возвращает **логическое** значение, указывающее, содержит ли указанной ячейке любого текста. Возвращает **значение True** , если указанный ячейки содержит текст. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **HasText**

 переменная _expression_A, представляет собой объект- **ячейки** .


## <a name="example"></a>Пример

Если фигура один по одному содержит таблицы и первой ячейки таблицы содержит текст, в этом примере отображается текст в окне сообщения.


```vb
With ActiveDocument.Pages(1).Shapes(1) 
 
 ' Check for table. 
 If .HasTable Then 
 With .Table.Cells(StartRow:=1, StartColumn:=1, _ 
 EndRow:=1, EndColumn:=1).Item(1) 
 
 ' Check for text in first cell. 
 If .HasText Then 
 MsgBox "Text from first cell of table: " _ 
 &; vbCr &; .Text 
 Else 
 MsgBox "No text in first cell." 
 End If 
 
 End With 
 Else 
 MsgBox "No table in shape one." 
 End If 
 
End With 

```


