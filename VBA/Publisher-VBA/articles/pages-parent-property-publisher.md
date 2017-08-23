---
title: "Свойство Pages.Parent (издатель)"
keywords: vbapb10.chm458754
f1_keywords: vbapb10.chm458754
ms.prod: publisher
api_name: Publisher.Pages.Parent
ms.assetid: 30de36f1-c65c-1549-c054-a20987a61ca7
ms.date: 06/08/2017
ms.openlocfilehash: 4ec0bbc6e8f0f6329b8880a00a85df82b7c6a37f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagesparent-property-publisher"></a>Свойство Pages.Parent (издатель)

Возвращает объект, представляющий родительский объект для указанного объекта. Например для объекта **[TextFrame](textframe-object-publisher.md)** возвращает объект **[фигуры](shape-object-publisher.md)** , представляющий родительскую фигуру рамки. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Родительский**

 переменная _expression_A, представляет собой объект- **страниц** .


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


