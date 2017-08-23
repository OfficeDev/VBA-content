---
title: "Свойство CalloutFormat.Parent (издатель)"
keywords: vbapb10.chm2490370
f1_keywords: vbapb10.chm2490370
ms.prod: publisher
api_name: Publisher.CalloutFormat.Parent
ms.assetid: 2cd29cfa-2def-b3cf-de50-948568f8968a
ms.date: 06/08/2017
ms.openlocfilehash: 81d628b2b340bbab7ab2d73d40d5cea70550e7a3
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="calloutformatparent-property-publisher"></a>Свойство CalloutFormat.Parent (издатель)

Возвращает объект, представляющий родительский объект для указанного объекта. Например для объекта **[TextFrame](textframe-object-publisher.md)** возвращает объект **[фигуры](shape-object-publisher.md)** , представляющий родительскую фигуру рамки. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Родительский**

 переменная _expression_A, представляет собой объект- **CalloutFormat** .


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


