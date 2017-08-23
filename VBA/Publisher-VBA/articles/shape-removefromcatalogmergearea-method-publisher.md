---
title: "Метод Shape.RemoveFromCatalogMergeArea (издатель)"
keywords: vbapb10.chm5308689
f1_keywords: vbapb10.chm5308689
ms.prod: publisher
api_name: Publisher.Shape.RemoveFromCatalogMergeArea
ms.assetid: 3b3630c3-6bf1-494b-151c-c930f32a2a77
ms.date: 06/08/2017
ms.openlocfilehash: 7369717a6ff5ad5b61ccebacb1464b8fd2fb92df
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperemovefromcatalogmergearea-method-publisher"></a>Метод Shape.RemoveFromCatalogMergeArea (издатель)

Удаляет из указанной странице области фигуры. Удаленные фигур не удаляются, но вместо этого остаются на месте на странице, содержащей области данных.


## <a name="syntax"></a>Синтаксис

 _выражение_. **RemoveFromCatalogMergeArea**

 переменная _expression_A, представляющий объект **фигуры** .


### <a name="return-value"></a>Возвращаемое значение

Значение Nothing


## <a name="remarks"></a>Заметки

Используйте метод **[AddToCatalogMergeArea](shape-addtocatalogmergearea-method-publisher.md)** объектов **[фигуры](shape-object-publisher.md)** или **[ShapeRange](shaperange-object-publisher.md)** добавление фигур в области объединения в каталог.

Метод **[RemoveCatalogMergeArea](shape-removecatalogmergearea-method-publisher.md)** объекта **[Shape](shape-object-publisher.md)** для удаления области данных со страницы публикации, но оставьте фигуры, которые он содержит.


## <a name="example"></a>Пример

Следующий пример проверяет ли любую страницу указанной публикации содержит область объединения в каталог. В случае любую страницу удаляются из области данных и удаления всех фигур и затем область данных удаляется из публикации.


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


