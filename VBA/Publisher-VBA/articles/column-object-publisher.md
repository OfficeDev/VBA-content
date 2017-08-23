---
title: "Объект столбца (издатель)"
keywords: vbapb10.chm5046271
f1_keywords: vbapb10.chm5046271
ms.prod: publisher
api_name: Publisher.Column
ms.assetid: 7f14fd4f-3919-8dd9-ed1e-988269b4b0c9
ms.date: 06/08/2017
ms.openlocfilehash: 1f415537f5bb718f50daeb0f068afce1bc92f019
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="column-object-publisher"></a>Объект столбца (издатель)

Представляет одного столбца. Объект **столбца** , является участником коллекции **[столбцов](columns-object-publisher.md)** . Коллекция **столбцов** включает все столбцы в таблице, выделения или диапазона.
 


## <a name="example"></a>Пример

Используйте **столбцы** (индекс), где индекс — номер столбца, чтобы возвратить объект одного **столбца** . Номер индекса представляет положение столбца в коллекции **столбцов** (начиная с слева направо). В этом примере выбирает три столбца в первой фигуры в активной публикации. В этом примере предполагается, что первой фигуры являются таблицы и не другого типа фигуры.
 

 

```
Sub SelectColumn() 
 ActiveDocument.Pages(2).Shapes(1).Table.Columns(3).Cells.Select 
End Sub
```

Метод **[Item](columns-item-method-publisher.md)** коллекции **[столбцов](columns-object-publisher.md)** для возврата объекта **столбца** . В этом примере вводит текст в первую ячейку третьего столбца указанную таблицу и форматирования текста с полужирный шрифт точка 15. В этом примере предполагается, что первой фигуры являются таблицы и не другого типа фигуры.
 

 



```
Sub ColumnHeading() 
 With ActiveDocument.Pages(2).Shapes(1).Table.Columns(3) _ 
 .Cells(1).Text 
 .Text = "Sales" 
 .Font.Bold = msoTrue 
 .Font.Size = 15 
 End With 
End Sub
```

Используйте метод **[Delete](column-delete-method-publisher.md)** для удаления столбца из таблицы. В этом примере удаляется столбцов, добавленных в приведенном выше примере.
 

 



```
Sub DeleteColumn() 
 ActiveDocument.Pages(2).Shapes(1).Table.Columns(3).Delete 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Delete](column-delete-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](column-application-property-publisher.md)|
|[Ячейки](column-cells-property-publisher.md)|
|[Родительский раздел](column-parent-property-publisher.md)|
|[Ширина](column-width-property-publisher.md)|

