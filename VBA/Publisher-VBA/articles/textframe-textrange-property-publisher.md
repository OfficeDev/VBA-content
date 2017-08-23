---
title: "Свойство TextFrame.TextRange (издатель)"
keywords: vbapb10.chm3866627
f1_keywords: vbapb10.chm3866627
ms.prod: publisher
api_name: Publisher.TextFrame.TextRange
ms.assetid: 44a8395e-81dc-7d06-f068-89f77a889f5e
ms.date: 06/08/2017
ms.openlocfilehash: d1e86568dcd5975f23d7822316060208cc81ceac
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textframetextrange-property-publisher"></a>Свойство TextFrame.TextRange (издатель)

Возвращает объект **[TextRange](textrange-object-publisher.md)** , представляющий текст, который присоединен к фигуре и свойства и методы для работы с текстом.


## <a name="syntax"></a>Синтаксис

 _выражение_. **TextRange**

 переменная _expression_A, представляет собой объект- **TextFrame** .


## <a name="example"></a>Пример

В следующем примере добавляется текст надписи фигуры один активный публикации и форматирует новый текст. В этом примере предполагается, что имеется по крайней мере один фигуры на первой странице active публикации.


```vb
Sub AddTextToTextFrame() 
 With ActiveDocument.Pages(1).TextFrame.TextRange 
 .Text = "My Text" 
 With .Font 
 .Bold = msoTrue 
 .Size = 25 
 .Name = "Arial" 
 End With 
 End With 
End Sub
```

В следующем примере добавляет прямоугольник active публикации и добавляет текст.




```vb
Sub AddTextToShape() 
 With ActiveDocument.Pages(1).Shapes.AddShape(Type:=msoShapeRectangle, _ 
 Left:=72, Top:=72, Width:=250, Height:=140) 
 .TextFrame.TextRange.Text = "Here is some test text" 
 End With 
End Sub
```


