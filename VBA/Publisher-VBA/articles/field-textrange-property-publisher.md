---
title: "Свойство Field.TextRange (издатель)"
keywords: vbapb10.chm6094852
f1_keywords: vbapb10.chm6094852
ms.prod: publisher
api_name: Publisher.Field.TextRange
ms.assetid: 09279cc7-3911-3b8d-51f2-b26494220c68
ms.date: 06/08/2017
ms.openlocfilehash: 8e1ace36cec87e2c6c790f87637a000d2e95d3a5
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fieldtextrange-property-publisher"></a>Свойство Field.TextRange (издатель)

Возвращает объект **[TextRange](textrange-object-publisher.md)** , представляющий текст, который присоединен к фигуры и свойства и методы для работы с текстом.


## <a name="syntax"></a>Синтаксис

 _выражение_. **TextRange**

 переменная _expression_A, представляющий объект **поля** .


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


