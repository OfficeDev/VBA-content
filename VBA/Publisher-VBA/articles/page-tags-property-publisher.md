---
title: "Свойство Page.Tags (издатель)"
keywords: vbapb10.chm393235
f1_keywords: vbapb10.chm393235
ms.prod: publisher
api_name: Publisher.Page.Tags
ms.assetid: 94a8be36-20c2-65bc-b1e2-41f24703b264
ms.date: 06/08/2017
ms.openlocfilehash: cb46a1cab2310ae7f98b4d9489f55790d82cac63
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagetags-property-publisher"></a>Свойство Page.Tags (издатель)

Возвращает коллекцию **[тегов](tags-object-publisher.md)** , представляющее теги или настраиваемых свойств, применяемых к фигуры, диапазона фигуры, страницы или публикации.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Теги**

 переменная _expression_A, представляющий объект **Page** .


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


