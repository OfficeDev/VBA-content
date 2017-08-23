---
title: "Метод Shape.RemoveCatalogMergeArea (издатель)"
keywords: vbapb10.chm5308691
f1_keywords: vbapb10.chm5308691
ms.prod: publisher
api_name: Publisher.Shape.RemoveCatalogMergeArea
ms.assetid: addff960-562e-b8e8-ec56-ddcf2b9ccaa7
ms.date: 06/08/2017
ms.openlocfilehash: c4077d31d0cfb6786e3f1edfb4decead517390c7
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperemovecatalogmergearea-method-publisher"></a>Метод Shape.RemoveCatalogMergeArea (издатель)

Удаление области данных из указанной публикации страницы. Все фигуры, содержащиеся в этой области остаются на месте на странице, но больше не подключены к каталогу источник данных.


## <a name="syntax"></a>Синтаксис

 _выражение_. **RemoveCatalogMergeArea**

 переменная _expression_A, представляющий объект **фигуры** .


## <a name="remarks"></a>Заметки

Удаление области объединения в каталог со страницы публикации не приводит к отключению источника данных из публикации. Используйте свойство **[IsDataSourceConnected](document-isdatasourceconnected-property-publisher.md)** объекта **[Document](document-object-publisher.md)** для определения, если источник данных подключается к публикации.

Метод **[AddCatalogMergeArea](shapes-addcatalogmergearea-method-publisher.md)** коллекцию **[фигур](shapes-object-publisher.md)** для добавления области объединения в каталог на публикацию. Страница публикации может содержать только одну область.


## <a name="example"></a>Пример

Следующий пример проверяет ли любую страницу в указанной публикации содержит область объединения в каталог. В случае любую страницу удаляются из области данных и удаления всех фигур и затем область данных удаляется из публикации.


```vb
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


