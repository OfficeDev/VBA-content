---
title: "Свойство FillFormat.BackColor (издатель)"
keywords: vbapb10.chm2359552
f1_keywords: vbapb10.chm2359552
ms.prod: publisher
api_name: Publisher.FillFormat.BackColor
ms.assetid: 61c6171b-f707-6741-68d2-5389bb3fac10
ms.date: 06/08/2017
ms.openlocfilehash: 6a71766c7831c34307bd0e3cd3c0127214e3d0ad
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fillformatbackcolor-property-publisher"></a>Свойство FillFormat.BackColor (издатель)

Возвращает или задает объект **[ColorFormat](colorformat-object-publisher.md)** , представляющее цвет фона для указанного заливка или узорная линия. .


## <a name="syntax"></a>Синтаксис

 _выражение_. **Цвет фона**

 переменная _expression_A, представляет собой объект- **FillFormat** .


## <a name="remarks"></a>Заметки

Свойство **[ForeColor](fillformat-forecolor-property-publisher.md)** задать цвет переднего плана для заполнения или строки.


## <a name="example"></a>Пример

В этом примере добавляет прямоугольник active публикации и затем задает цвет переднего плана, цвет фона и градиент для заливки прямоугольника.


```vb
With ActiveDocument.Pages(1).Shapes.AddShape _ 
 (Type:=msoShapeRectangle, _ 
 Left:=90, Top:=90, Width:=90, Height:=50).Fill 
 .ForeColor.RGB = RGB(128, 0, 0) 
 .BackColor.RGB = RGB(170, 170, 170) 
 .TwoColorGradient _ 
 Style:=msoGradientHorizontal, Variant:=1 
End With
```

В этом примере добавляется узорная линия active публикации.




```vb
With ActiveDocument.Pages(1).Shapes.AddLine _ 
 (BeginX:=10, BeginY:=100, EndX:=250, EndY:=0).Line 
 .Weight = 6 
 .ForeColor.RGB = RGB(0, 0, 255) 
 .BackColor.RGB = RGB(128, 0, 0) 
 .Pattern = msoPatternDarkDownwardDiagonal 
End With 

```


