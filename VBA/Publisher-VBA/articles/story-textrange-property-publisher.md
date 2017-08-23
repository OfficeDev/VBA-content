---
title: "Свойство Story.TextRange (издатель)"
keywords: vbapb10.chm5832712
f1_keywords: vbapb10.chm5832712
ms.prod: publisher
api_name: Publisher.Story.TextRange
ms.assetid: c948da79-ea67-0c8c-1df3-2b32499ea9b3
ms.date: 06/08/2017
ms.openlocfilehash: 04193d58adc4c0fe61b96c82534d208f6fd799f0
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="storytextrange-property-publisher"></a>Свойство Story.TextRange (издатель)

Возвращает объект **[TextRange](textrange-object-publisher.md)** , представляющий текст, который присоединен к фигуры и свойства и методы для работы с текстом.


## <a name="syntax"></a>Синтаксис

 _выражение_. **TextRange**

 переменная _expression_A, представляет собой объект- **материала** .


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


