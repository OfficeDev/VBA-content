---
title: "Свойство Shape.Shadow (издатель)"
keywords: vbapb10.chm2228296
f1_keywords: vbapb10.chm2228296
ms.prod: publisher
api_name: Publisher.Shape.Shadow
ms.assetid: cfb908ae-ef1d-9539-1f82-2693cbe38d97
ms.date: 06/08/2017
ms.openlocfilehash: cde9675f703c70e4e31764461b09a8a408d35866
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapeshadow-property-publisher"></a>Свойство Shape.Shadow (издатель)

Возвращает объект **[ShadowFormat](shadowformat-object-publisher.md)** , представляющий тени для указанной фигуры.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Теневая**

 переменная _expression_A, представляющий объект **фигуры** .


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


