---
title: "Объект строк (издатель)"
keywords: vbapb10.chm4980735
f1_keywords: vbapb10.chm4980735
ms.prod: publisher
api_name: Publisher.Rows
ms.assetid: 31b04a41-9005-8f51-87ab-426af0e901ed
ms.date: 06/08/2017
ms.openlocfilehash: 542222825e358c9d261346f650d7061548aff9f2
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="rows-object-publisher"></a>Объект строк (издатель)

Коллекция объектов **[строк](row-object-publisher.md)** , представляющих строки в таблице.
 


## <a name="example"></a>Пример

Свойство **[строки](table-rows-property-publisher.md)** **[в таблице](table-object-publisher.md)** объектов для возврата коллекции **строк** . Следующий пример показывает число объекты **[строк](row-object-publisher.md)** в коллекции **строк** для первой таблицы в активный документ.
 

 

```
Sub CountRows() 
 MsgBox ActiveDocument.Pages(2).Shapes(1).Table.Rows.Count 
End Sub
```

В этом примере задается заливки для всех четных строк и очищает заливки для всех нечетных строк в указанной таблице. В этом примере предполагается, что указанные форму — это таблица и не другого типа фигуры.
 

 



```
Sub FillCellsByRow() 
 Dim shpTable As Shape 
 Dim rowTable As Row 
 Dim celTable As Cell 
 
 Set shpTable = ActiveDocument.Pages(2).Shapes(1) 
 For Each rowTable In shpTable.Table.Rows 
 For Each celTable In rowTable.Cells 
 If celTable.Row Mod 2 = 0 Then 
 celTable.Fill.ForeColor.RGB = RGB _ 
 (Red:=180, Green:=180, Blue:=180) 
 Else 
 celTable.Fill.ForeColor.RGB = RGB _ 
 (Red:=255, Green:=255, Blue:=255) 
 End If 
 Next celTable 
 Next rowTable 
 
End Sub
```

Использование **строк** (индекс), где индекс — номер индекса, для возврата объекта **строки** . Номер индекса представляет положение строки в коллекции **строк** (начиная с слева направо). В следующем примере выбирается третьей строки в указанной таблице.
 

 



```
Sub SelectRows() 
 ActiveDocument.Pages(2).Shapes(1).Table.Rows(3).Cells.Select 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Добавление](rows-add-method-publisher.md)|
|[Элемент](rows-item-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](rows-application-property-publisher.md)|
|[Count](rows-count-property-publisher.md)|
|[Родительский раздел](rows-parent-property-publisher.md)|

