---
title: "Объект CellRange (издатель)"
keywords: vbapb10.chm5242879
f1_keywords: vbapb10.chm5242879
ms.prod: publisher
api_name: Publisher.CellRange
ms.assetid: 86e164f3-2a04-013f-3da8-d45c013eae7b
ms.date: 06/08/2017
ms.openlocfilehash: 03d5c8f5b8744d0fb20fbd576bb2f1f6dc9874a9
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="cellrange-object-publisher"></a>Объект CellRange (издатель)

Коллекция объектов **[ячейки](cell-object-publisher.md)** в строку или столбец таблицы. Коллекция **CellRange** представляет всем ячейкам в строку или указанный столбец.
 


## <a name="remarks"></a>Заметки

Несмотря на то, что объект коллекции с именем **CellRange** и отображаются в обозревателе объектов, это ключевое слово не используется в программирование объектной модели Microsoft Publisher. Вместо этого используется ключевое слово **ячеек** .
 

 
Не удается программными средствами добавить или удалить отдельные ячейки из таблицы Publisher. Использование метода **[AddTable](shapes-addtable-method-publisher.md)** с коллекцию **[фигур](shapes-object-publisher.md)** для добавления новой таблицы в публикацию. Используйте метод **[Add](columns-add-method-publisher.md)** коллекции **[столбцы](columns-object-publisher.md)** или **[строки](rows-object-publisher.md)** для добавления столбца или строки в таблицу. Используйте метод **[Delete](column-delete-method-publisher.md)** **столбцы** или **строки** семейств для удаления столбца или строки из таблицы.
 

 

## <a name="example"></a>Пример

Свойство **[ячейки](column-cells-property-publisher.md)** используется для возврата коллекции **CellRange** . В этом примере выполняется объединение ячеек в первый столбец таблицы.
 

 

```
Sub MergeCellsInFirstColumn() 
 With ActiveDocument.Pages(1).Shapes(1).Table 
 .Cells(StartRow:=1, StartColumn:=1, _ 
 EndRow:=.Rows.Count, EndColumn:=1).Select 
 End With 
 Selection.TableCellRange.Merge 
End Sub
```

Свойство **[Count](cellrange-count-property-publisher.md)** возвращает число ячеек в строки, столбца, таблицы или выбора. В этом примере выводится сообщение с количеством ячеек указанную таблицу.
 

 



```
Sub NumberOfTableCells() 
 MsgBox ActiveDocument.Pages(1).Shapes(1).Table _ 
 .Cells.Count 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Элемент](cellrange-item-method-publisher.md)|
|[Объединение](cellrange-merge-method-publisher.md)|
|[Выберите](cellrange-select-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](cellrange-application-property-publisher.md)|
|[Столбец](cellrange-column-property-publisher.md)|
|[Count](cellrange-count-property-publisher.md)|
|[Высота](cellrange-height-property-publisher.md)|
|[Родительский раздел](cellrange-parent-property-publisher.md)|
|[Строка](cellrange-row-property-publisher.md)|
|[Ширина](cellrange-width-property-publisher.md)|

