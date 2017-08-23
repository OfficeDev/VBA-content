---
title: "Свойство ShadowFormat.ForeColor (издатель)"
keywords: vbapb10.chm3670272
f1_keywords: vbapb10.chm3670272
ms.prod: publisher
api_name: Publisher.ShadowFormat.ForeColor
ms.assetid: 1ff2210f-1ab4-e991-746b-d4383a87c9e8
ms.date: 06/08/2017
ms.openlocfilehash: fe330377e6c80ead444aa2f31ef701b3b13785bc
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shadowformatforecolor-property-publisher"></a>Свойство ShadowFormat.ForeColor (издатель)

Возвращает или задает объект **[ColorFormat](colorformat-object-publisher.md)** , представляющее цвет переднего плана для заливки, строки или тени. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Цвет текста**

 переменная _expression_A, представляет собой объект- **ShadowFormat** .


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


