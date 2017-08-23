---
title: "Свойство Shape.ParentGroupShape (издатель)"
keywords: vbapb10.chm2228338
f1_keywords: vbapb10.chm2228338
ms.prod: publisher
api_name: Publisher.Shape.ParentGroupShape
ms.assetid: ced4c348-4ef5-c703-fdea-65c33d37b4c0
ms.date: 06/08/2017
ms.openlocfilehash: a322eae8794503f3e3d5d1fda64bf5c4c3faef17
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapeparentgroupshape-property-publisher"></a>Свойство Shape.ParentGroupShape (издатель)

Возвращает объект **[фигуры](shape-object-publisher.md)** , представляющий распространенных родительскую фигуру фигуры дочерних или диапазона фигуры потомков.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ParentGroupShape**

 переменная _expression_A, представляющий объект **фигуры** .


### <a name="return-value"></a>Возвращаемое значение

Shape


## <a name="example"></a>Пример

В этом примере создается две фигуры в активный документ и групп этих фигур. Затем с помощью одной фигуры в группе, она получает доступ к родительской группы и всех фигур в родительской группы с той же схеме заливки. В этом примере предполагается, что активный документ в настоящий момент нет фигуры. В этом случае может возникнуть ошибка.


```vb
Sub ParentGroupShape() 
 Dim shpGroup As Shape 
 
 With ActiveDocument.Pages(1).Shapes 
 .AddShape Type:=msoShapeOval, Left:=72, _ 
 Top:=72, Width:=100, Height:=100 
 .AddShape Type:=msoShapeHeart, Left:=110, _ 
 Top:=120, Width:=100, Height:=100 
 .Range(Array(1, 2)).Group 
 End With 
 
 Set shpGroup = ActiveDocument.Pages(1).Shapes(1) _ 
 .GroupItems(1).ParentGroupShape 
 shpGroup.Fill.Patterned Pattern:=msoPattern25Percent 
 
End Sub
```


