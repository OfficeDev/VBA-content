---
title: "Объект строки (издатель)"
keywords: vbapb10.chm4915199
f1_keywords: vbapb10.chm4915199
ms.prod: publisher
api_name: Publisher.Row
ms.assetid: 11f4688b-b94e-fa09-7c1b-4cbcca330936
ms.date: 06/08/2017
ms.openlocfilehash: 84ab4acc3e551d977a5a08480099db71482f7f5e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="row-object-publisher"></a>Объект строки (издатель)

Представляет строку в таблице. Объект **строки** является членом коллекции **[строк](rows-object-publisher.md)** . Набор **строк** включает в себя все строки в указанной таблице.
 


## <a name="example"></a>Пример

Использование **строк** (индекс), где индекс — число строк, для возврата объекта **строки** . Номер индекса представляет положение строки в коллекции **строк** (начиная с слева направо). В этом примере выделяет первую строку в первую фигуру на второй active публикации. В этом примере предполагается, что указанные форму — это таблица и не другого типа фигуры.
 

 

```
Sub SelectRow() 
 ActiveDocument.Pages(2).Shapes(1).Table.Rows(1).Cells.Select 
End Sub
```

Метод **[Item](rows-item-method-publisher.md)** коллекции **[строк](rows-object-publisher.md)** для возврата объекта **строки** . В этом примере задается заливки для всех четных строк и очищает заливки для всех нечетной нумерованный строк в указанной таблице. В этом примере предполагается, что указанные форму — это таблица и не другого типа фигуры.
 

 



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

Используйте метод **[Add](rows-add-method-publisher.md)** для добавления строки в таблицу. В этом примере добавляет строку в указанную таблицу на второй странице active публикации и затем изменяет ширину, выполняется объединение ячеек и задает цвет заливки. В этом примере предполагается, что первой фигуры являются таблицы и не другого типа фигуры.
 

 



```
Sub NewRow() 
 Dim rowNew As Row 
 
 Set rowNew = ActiveDocument.Pages(2).Shapes(1).Table.Rows _ 
 .Add(BeforeRow:=3) 
 With rowNew 
 .Height = 2 
 .Cells.Merge 
 .Cells(1).Fill.ForeColor.RGB = RGB(Red:=0, Green:=0, Blue:=0) 
 End With 
End Sub
```

Используйте метод **[Delete](row-delete-method-publisher.md)** , чтобы удалить строку из таблицы. В этом примере удаляется добавленной в приведенном выше примере строки.
 

 



```
Sub DeleteRow() 
 ActiveDocument.Pages(2).Shapes(1).Table.Rows(3).Delete 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Delete](row-delete-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](row-application-property-publisher.md)|
|[Ячейки](row-cells-property-publisher.md)|
|[Высота](row-height-property-publisher.md)|
|[Родительский раздел](row-parent-property-publisher.md)|

