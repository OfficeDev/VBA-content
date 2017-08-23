---
title: "Объект CatalogMergeShapes (издатель)"
keywords: vbapb10.chm8454143
f1_keywords: vbapb10.chm8454143
ms.prod: publisher
api_name: Publisher.CatalogMergeShapes
ms.assetid: 1108e9a4-57ef-2b1a-0998-54b6fad838da
ms.date: 06/08/2017
ms.openlocfilehash: 5b9180b10cd1fb019be265c69e555bac423e66eb
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="catalogmergeshapes-object-publisher"></a>Объект CatalogMergeShapes (издатель)

Представляет фигуры, содержащиеся в этой области указанной публикации.
 


## <a name="remarks"></a>Заметки

Область данных автоматически изменяется, чтобы вместить объекты, размер которых затем области объединения или, находятся вне области данных после их добавления.
 

 
Фигуры внутри области данных автоматически изменения размера или положения, если область данных уменьшается размер или перемещены.
 

 
Область данных может содержать изображения и текст полей данных, вставленных, в того в другие элементы дизайна, выбранное. 
 

 

## <a name="example"></a>Пример

Свойство **[CatalogMergeItems](shape-catalogmergeitems-property-publisher.md)** **[фигуры](shape-object-publisher.md)** или **[ShapeRange](shaperange-object-publisher.md)** объекты для возвращения содержимого области данных. Следующий пример проверяет, является ли указанный публикация содержит область объединения в каталог. Если это так, возвращает список фигуры, которые он содержит.
 

 

```
Sub ListCatalogMergeAreaContents() 
 
 Dim pgPage As Page 
 Dim mmLoop As Shape 
 Dim intCount As Integer 
 
 For Each pgPage In ThisDocument.Pages 
 For Each mmLoop In pgPage.Shapes 
 
 If mmLoop.Type = pbCatalogMergeArea Then 
 
 With mmLoop.CatalogMergeItems 
 For intCount = 1 To .Count 
 Debug.Print "Shape ID: " &amp; _ 
 mmLoop.CatalogMergeItems.Item(intCount).ID 
 Debug.Print "Shape Name: " &amp; _ 
 mmLoop.CatalogMergeItems.Item(intCount).Name 
 Next 
 End With 
 
 End If 
 
 Next mmLoop 
 Next pgPage 
 
End Sub 

```

Используйте метод **[AddToCatalogMergeArea](shape-addtocatalogmergearea-method-publisher.md)** объектов **[фигуры](shape-object-publisher.md)** или **[ShapeRange](shaperange-object-publisher.md)** добавление фигур в области объединения в каталог. Следующий пример добавляет прямоугольник области данных в указанной публикации. В этом примере предполагается, что область объединения в каталог был добавлен к первой страницы публикации.
 

 



```
ThisDocument.Pages(1).Shapes.AddShape(1, 80, 75, 450, 125).AddToCatalogMergeArea
```

Используйте **CatalogMergeItems** (индекс), где индекс — номер индекса, для возвращения фигуры области объединения один каталог. Следующий пример удаляет первую фигуру из области данных.
 

 



```
ThisDocument.Pages(1).Shapes(1).CatalogMergeItems(1).RemoveFromCatalogMergeArea
```

Метод **[RemoveFromCatalogMergeArea](shape-removefromcatalogmergearea-method-publisher.md)** **[фигуры](shape-object-publisher.md)** или **[ShapeRange](shaperange-object-publisher.md)** объектов для удаления из области фигуры. Удаленные фигур не удаляются, но вместо этого помещаются на странице публикации, содержащий области данных. Следующий пример проверяет, является ли указанный публикация содержит область объединения в каталог. Если это так, удаляются из области данных и удаления всех фигур и затем область данных удаляется из публикации.
 

 



```
Sub DeleteCatalogMergeAreaAndAllShapesWithin() 
 Dim pgPage As Page 
 Dim mmLoop As Shape 
 Dim intCount As Integer 
 Dim strName As String 
 
 For Each pgPage In ThisDocument.Pages 
 For Each mmLoop In pgPage.Shapes 
 
 If mmLoop.Type = pbCatalogMergeArea Then 
 With mmLoop.CatalogMergeItems 
 For intCount = .Count To 1 Step -1 
 strName = mmLoop.CatalogMergeItems.Item(intCount).Name 
 .Item(intCount).RemoveFromCatalogMergeArea 
 pgPage.Shapes(strName).Delete 
 Next 
 End With 
 mmLoop.RemoveCatalogMergeArea 
 End If 
 
 Next mmLoop 
 Next pgPage 
 
 End Sub 

```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Элемент](catalogmergeshapes-item-method-publisher.md)|
|[Range](catalogmergeshapes-range-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](catalogmergeshapes-application-property-publisher.md)|
|[Count](catalogmergeshapes-count-property-publisher.md)|
|[HorizontalRepeat](catalogmergeshapes-horizontalrepeat-property-publisher.md)|
|[Родительский раздел](catalogmergeshapes-parent-property-publisher.md)|
|[VerticalRepeat](catalogmergeshapes-verticalrepeat-property-publisher.md)|

