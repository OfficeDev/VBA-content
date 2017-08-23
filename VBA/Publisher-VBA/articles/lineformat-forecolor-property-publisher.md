---
title: "Свойство LineFormat.ForeColor (издатель)"
keywords: vbapb10.chm3408136
f1_keywords: vbapb10.chm3408136
ms.prod: publisher
api_name: Publisher.LineFormat.ForeColor
ms.assetid: 192314ba-dbca-cce0-25c4-6e276a4f268b
ms.date: 06/08/2017
ms.openlocfilehash: d1d8400d613f8cb87167890380bf9ff3d0bf5e3a
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="lineformatforecolor-property-publisher"></a>Свойство LineFormat.ForeColor (издатель)

Возвращает или задает объект **[ColorFormat](colorformat-object-publisher.md)** , представляющее цвет переднего плана для заливки, строки или тени. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Цвет текста**

 переменная _expression_A, представляет собой объект- **LineFormat** .


## <a name="remarks"></a>Заметки

Свойство **BackColor** задайте цвет фона для заполнения или строку.


## <a name="example"></a>Пример

В этом примере добавляет прямоугольник active публикации и затем задает цвет переднего плана, цвет фона и градиент для заливки прямоугольника.


```vb
With ActiveDocument.Pages(1).Shapes.AddShape _ 
 (Type:=msoShapeRectangle, _ 
 Left:=90, Top:=90, Width:=90, Height:=50).Fill 
 .ForeColor.RGB = RGB(128, 0, 0) 
 .BackColor.RGB = RGB(170, 170, 170) 
 .TwoColorGradient msoGradientHorizontal, 1 
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


