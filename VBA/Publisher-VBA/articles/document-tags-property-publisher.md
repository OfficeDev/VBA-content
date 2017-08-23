---
title: "Свойство Document.Tags (издатель)"
keywords: vbapb10.chm196661
f1_keywords: vbapb10.chm196661
ms.prod: publisher
api_name: Publisher.Document.Tags
ms.assetid: d8baaf50-86ad-1997-c1b3-e54a77a3ee5b
ms.date: 06/08/2017
ms.openlocfilehash: 3a52fafd2dc5b7077fa1dcd24c4ac2cfcd3a59db
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documenttags-property-publisher"></a>Свойство Document.Tags (издатель)

Возвращает коллекцию **[тегов](tags-object-publisher.md)** , представляющее теги или настраиваемых свойств, применяемых к фигуры, диапазона фигуры, страницы или публикации.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Теги**

 переменная _expression_A, представляющий объект **Document** .


## <a name="example"></a>Пример

В этом примере добавляется тег для каждой фигуры овала на первой странице active публикации.


```vb
Dim shp As Shape 
Dim tagsAll As Tags 
Dim tagLoop As Tag 
Dim blnTag As Boolean 
 
With ActiveDocument.Pages(1) 
 For Each shp In .Shapes 
 If shp.AutoShapeType = msoShapeOval Then 
 Set tagsAll = shp.Tags 
 blnTag = False 
 
 For Each tagLoop In tagsAll 
 If tagLoop.Name = "Shape" Then 
 blnTag = True 
 Exit For 
 End If 
 Next tagLoop 
 
 If blnTag = False Then 
 tagsAll.Add Name:="Shape", Value:="Oval" 
 End If 
 End If 
 Next shp 
End With 

```


