---
title: "Свойство Shape.Tags (издатель)"
keywords: vbapb10.chm2228329
f1_keywords: vbapb10.chm2228329
ms.prod: publisher
api_name: Publisher.Shape.Tags
ms.assetid: 282f77c8-f075-1eeb-65e8-f1126def32ff
ms.date: 06/08/2017
ms.openlocfilehash: 78b5b7e0a9092df39de9a89861c9edbbfeddaf8b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapetags-property-publisher"></a>Свойство Shape.Tags (издатель)

Возвращает коллекцию **[тегов](tags-object-publisher.md)** , представляющее теги или настраиваемых свойств, применяемых к фигуры, диапазона фигуры, страницы или публикации.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Теги**

 переменная _expression_A, представляющий объект **фигуры** .


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


