---
title: "Объект столбцы (издатель)"
keywords: vbapb10.chm5111807
f1_keywords: vbapb10.chm5111807
ms.prod: publisher
api_name: Publisher.Columns
ms.assetid: 3fe6ddce-a598-a967-fc89-7296c18a6a55
ms.date: 06/08/2017
ms.openlocfilehash: a586b572a5e50afc1af71cfb63fcd3d49017430b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="columns-object-publisher"></a>Объект столбцы (издатель)

Коллекция объектов **[столбцов](column-object-publisher.md)** , которые представляют столбцы в таблице.
 


## <a name="example"></a>Пример

Свойство **[столбцы](table-columns-property-publisher.md)** объекта **[таблицы](table-object-publisher.md)** для возврата коллекции **столбцов** . Следующий пример показывает число объектов **[столбца](column-object-publisher.md)** в коллекции **столбцов** для первой таблицы в активный документ.
 

 

```
Sub CountColumns() 
 MsgBox "The number of columns in the table is " &amp; _ 
 ActiveDocument.Pages(2).Shapes(1).Table.Columns.Count 
End Sub
```

В этом примере вводит полужирный номер в каждой ячейки в указанной таблице. В этом примере предполагается, что указанные форму — это таблица и не другого типа фигуры.
 

 



```
Sub CountCellsByColumn() 
 Dim shpTable As Shape 
 Dim colTable As Column 
 Dim celTable As Cell 
 Dim intCount As Integer 
 
 intCount = 1 
 
 Set shpTable = ActiveDocument.Pages(2).Shapes(1) 
 For Each colTable In shpTable.Table.Columns 
 For Each celTable In colTable.Cells 
 With celTable.Text 
 .Text = intCount 
 .ParagraphFormat.Alignment = _ 
 pbParagraphAlignmentCenter 
 .Font.Bold = msoTrue 
 intCount = intCount + 1 
 End With 
 Next celTable 
 Next colTable 
 
End Sub
```

Используйте **столбцы** (индекс), где индекс — номер индекса, чтобы возвратить объект одного **столбца** . Номер индекса представляет положение столбца в коллекции **столбцов** (начиная с слева направо). В следующем примере выбирается третьего столбца в указанной таблице.
 

 



```
Sub SelectColumns() 
 ActiveDocument.Pages(2).Shapes(1).Table.Columns(3).Cells.Select 
End Sub
```

Используйте метод **[Add](columns-add-method-publisher.md)** для добавления столбца в таблицу. В этом примере добавляет столбец в указанную таблицу на второй странице active публикации и затем изменяет ширину, слияния ячеек и задает цвет заливки. В этом примере предполагается, что первой фигуры являются таблицы и не другого типа фигуры.
 

 



```
Sub NewColumn() 
 Dim colNew As Column 
 
 Set colNew = ActiveDocument.Pages(2).Shapes(1).Table.Columns _ 
 .Add(BeforeColumn:=3) 
 With colNew 
 .Width = 2 
 .Cells.Merge 
 .Cells(1).Fill.ForeColor.RGB = RGB(Red:=202, Green:=202, Blue:=202) 
 End With 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Добавление](columns-add-method-publisher.md)|
|[Элемент](columns-item-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](columns-application-property-publisher.md)|
|[Count](columns-count-property-publisher.md)|
|[Родительский раздел](columns-parent-property-publisher.md)|

