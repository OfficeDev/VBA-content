---
title: "Свойство TextEffectFormat.Parent (издатель)"
keywords: vbapb10.chm3735554
f1_keywords: vbapb10.chm3735554
ms.prod: publisher
api_name: Publisher.TextEffectFormat.Parent
ms.assetid: 91e31549-b441-8cf3-6191-0f7adbc85cc1
ms.date: 06/08/2017
ms.openlocfilehash: e01ab7f31a2df3aaff0fa35142ae5308b2f971c8
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="texteffectformatparent-property-publisher"></a>Свойство TextEffectFormat.Parent (издатель)

Возвращает объект, представляющий родительский объект для указанного объекта. Например для объекта **[TextFrame](textframe-object-publisher.md)** возвращает объект **[фигуры](shape-object-publisher.md)** , представляющий родительскую фигуру рамки. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Родительский**

 переменная _expression_A, представляет собой объект- **TextEffectFormat** .


## <a name="example"></a>Пример

В этом примере обращается к родительский объект выбранной фигуры и добавляет новую форму и задает заливки для новой фигуры.


```vb
Sub ParentObject() 
 Dim shp As Shape 
 Dim pg As Page 
 
 Set pg = Selection.ShapeRange(1).Parent 
 Set shp = pg.Shapes.AddShape(Type:=msoShape5pointStar, _ 
 Left:=72, Top:=72, Width:=72, Height:=72) 
 
 shp.Fill.ForeColor.RGB = RGB(Red:=180, Green:=180, Blue:=180) 
End Sub
```

В этом примере возвращает родительский объект frame текст является первой фигуры в активной публикации, а затем заполняет фигуры с шаблоном.




```vb
Sub ParentShape() 
 Dim shpParent As Shape 
 Set shpParent = ActiveDocument.Pages(1).Shapes(1).TextFrame.Parent 
 shpParent.Fill.Patterned Pattern:=msoPatternSphere 
End Sub
```


