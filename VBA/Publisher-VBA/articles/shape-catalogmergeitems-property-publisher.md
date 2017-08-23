---
title: "Свойство Shape.CatalogMergeItems (издатель)"
keywords: vbapb10.chm5308690
f1_keywords: vbapb10.chm5308690
ms.prod: publisher
api_name: Publisher.Shape.CatalogMergeItems
ms.assetid: 1dcf4ae0-7a18-f1d5-2176-1912c63eefcc
ms.date: 06/08/2017
ms.openlocfilehash: 9b7fe17ec6791542e85e2613eb49075049cb8eeb
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapecatalogmergeitems-property-publisher"></a>Свойство Shape.CatalogMergeItems (издатель)

Возвращает коллекцию **CatalogMergeShapes** , представляющий фигур, включенные в этой области. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **CatalogMergeItems**

 переменная _expression_A, представляющий объект **фигуры** .


### <a name="return-value"></a>Возвращаемое значение

CatalogMergeShapes


## <a name="remarks"></a>Заметки

Область данных может содержать изображения и текст полей данных, вставленных, помимо другие элементы дизайна, выбранное.


## <a name="example"></a>Пример

Следующий пример проверяет ли любую страницу в указанной публикации содержит область объединения в каталог, и, если это так, он возвращает список фигур, которые он содержит.


```vb
Sub ListCatalogMergeAreaContents() 
 
 Dim pgPage As Page 
 Dim mmLoop As Shape 
 Dim intCount As Integer 
 
 For Each pgPage In ThisDocument.Pages 
 For Each mmLoop In pgPage.Shapes 
 
 If mmLoop.Type = pbCatalogMergeArea Then 
 
 With mmLoop.CatalogMergeItems 
 For intCount = 1 To .Count 
 Debug.Print "Shape ID: " &; _ 
 mmLoop.CatalogMergeItems.Item(intCount).ID 
 Debug.Print "Shape Name: " &; _ 
 mmLoop.CatalogMergeItems.Item(intCount).Name 
 Next 
 End With 
 
 End If 
 
 Next mmLoop 
 Next pgPage 
 
End Sub
```


