---
title: "Свойство ShapeRange.TextFrame (издатель)"
keywords: vbapb10.chm2293840
f1_keywords: vbapb10.chm2293840
ms.prod: publisher
api_name: Publisher.ShapeRange.TextFrame
ms.assetid: 2dbb7fb4-3ae4-d4c1-8b7e-3e087e32a96f
ms.date: 06/08/2017
ms.openlocfilehash: 4a7e7a51889e1faa5ba3552f504921618d6fb1c9
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangetextframe-property-publisher"></a>Свойство ShapeRange.TextFrame (издатель)

Возвращает объект **[TextFrame](textframe-object-publisher.md)** , представляющий текст в фигуру и свойства, которые управляют поля и ориентацию текста.


## <a name="syntax"></a>Синтаксис

 _выражение_. **TextFrame**

 переменная _expression_A, представляющий объект **ShapeRange** .


## <a name="example"></a>Пример

В следующем примере добавляется текст надписи фигуры один активный публикации и форматирует новый текст. В этом примере предполагается, что имеется по крайней мере один фигуры на первой странице active публикации.


```vb
Sub AddTextToTextFrame() 
 With ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange 
 .Text = "My Text" 
 With .Font 
 .Bold = msoTrue 
 .Size = 25 
 .Name = "Arial" 
 End With 
 End With 
End Sub
```


