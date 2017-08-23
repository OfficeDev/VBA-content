---
title: "Свойство ShapeRange.Tags (издатель)"
keywords: vbapb10.chm2293865
f1_keywords: vbapb10.chm2293865
ms.prod: publisher
api_name: Publisher.ShapeRange.Tags
ms.assetid: 792e3505-2c40-26e7-53c6-d50d84df22bb
ms.date: 06/08/2017
ms.openlocfilehash: c0b88f196baa94b747e34e1d7225b818923bc314
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangetags-property-publisher"></a>Свойство ShapeRange.Tags (издатель)

Возвращает коллекцию **[тегов](tags-object-publisher.md)** , представляющее теги или настраиваемых свойств, применяемых к фигуры, диапазона фигуры, страницы или публикации.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Теги**

 переменная _expression_A, представляющий объект **ShapeRange** .


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


