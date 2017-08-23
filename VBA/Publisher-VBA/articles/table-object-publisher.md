---
title: "Объект таблицы (издатель)"
keywords: vbapb10.chm4849663
f1_keywords: vbapb10.chm4849663
ms.prod: publisher
api_name: Publisher.Table
ms.assetid: 09da4a0a-2230-067e-1cac-55321ea044c5
ms.date: 06/08/2017
ms.openlocfilehash: 935e4927c2611cc4f98aa9e37fdbded767274f0d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="table-object-publisher"></a>Объект таблицы (издатель)

Представляет одну таблицу.


## <a name="example"></a>Пример

Используйте свойство **[таблицы](http://msdn.microsoft.com/library/a9b29d1f-2459-556c-56f8-f8f809b879c9%28Office.15%29.aspx)** для возврата объекта **в таблице** . В следующем примере выбирается указанную таблицу в активной публикации.


```
Sub SelectTable() 
 With ActiveDocument.Pages(1).Shapes(1) 
 If .Type = pbTable Then _ 
 .Table.Cells.Select 
 End With 
End Sub
```

Метод **[AddTable](http://msdn.microsoft.com/library/1aa00f40-de41-12ed-8d4f-5e9c91cbf5af%28Office.15%29.aspx)** используется для добавления объекта **Shape** , представляющее таблицы в указанном диапазоне. В следующем примере добавляется таблица 5 x 5 на первой странице active публикации и затем выбирает первый столбец новой таблицы.




```
Sub NewTable() 
 With ActiveDocument.Pages(1).Shapes.AddTable(NumRows:=5, NumColumns:=5, _ 
 Left:=72, Top:=300, Width:=400, Height:=100) 
 .Table.Columns(1).Cells.Select 
 End With 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[ApplyAutoFormat](http://msdn.microsoft.com/library/f792a5f3-0d1c-06de-a030-7a588ca372d2%28Office.15%29.aspx)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](http://msdn.microsoft.com/library/9d808ec1-3f29-c2d4-b685-7acd3c6d0f18%28Office.15%29.aspx)|
|[Ячейки](http://msdn.microsoft.com/library/42622697-aef1-0765-7d85-4919c298d92f%28Office.15%29.aspx)|
|[Столбцы](http://msdn.microsoft.com/library/fb55ba62-64a4-2221-3cc7-b349dc2f6934%28Office.15%29.aspx)|
|[GrowToFitText](http://msdn.microsoft.com/library/d8822df7-a252-a5bb-be26-83df8ec5eb94%28Office.15%29.aspx)|
|[Родительский раздел](http://msdn.microsoft.com/library/e7c02be8-1888-4817-05bf-75b030e597fc%28Office.15%29.aspx)|
|[Строк](http://msdn.microsoft.com/library/97a543b9-a1d7-c7f8-9f3c-e08256e0b364%28Office.15%29.aspx)|
|[TableDirection](http://msdn.microsoft.com/library/ffd664a8-781f-8fdc-055c-1ea7309b3b38%28Office.15%29.aspx)|

