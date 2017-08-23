---
title: "Свойство ShapeRange.Shadow (издатель)"
keywords: vbapb10.chm2293832
f1_keywords: vbapb10.chm2293832
ms.prod: publisher
api_name: Publisher.ShapeRange.Shadow
ms.assetid: d6ee257c-9a26-abfc-9e8e-ef89bf627690
ms.date: 06/08/2017
ms.openlocfilehash: 2ecd6fbab45dd68cf1539b47a1ca7296cff9edf7
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangeshadow-property-publisher"></a>Свойство ShapeRange.Shadow (издатель)

Возвращает объект **[ShadowFormat](shadowformat-object-publisher.md)** , представляющий тени для указанной фигуры.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Теневая**

 переменная _expression_A, представляющий объект **ShapeRange** .


## <a name="example"></a>Пример

В этом примере добавляется Стрелка с форматированием и заливки цвет затенения для первой страницы в активном документе.


```vb
Sub SetShapeShadow() 
 With ActiveDocument.Pages(1).Shapes.AddShape( _ 
 Type:=msoShapeRightArrow, Left:=72, _ 
 Top:=72, Width:=64, Height:=43) 
 .Shadow.Type = msoShadow5 
 .Fill.ForeColor.RGB = RGB(Red:=255, Green:=0, Blue:=255) 
 End With 
End Sub
```


