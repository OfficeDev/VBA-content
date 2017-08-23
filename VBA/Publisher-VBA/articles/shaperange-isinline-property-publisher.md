---
title: "Свойство ShapeRange.IsInline (издатель)"
keywords: vbapb10.chm2294022
f1_keywords: vbapb10.chm2294022
ms.prod: publisher
api_name: Publisher.ShapeRange.IsInline
ms.assetid: 32e038cc-5837-93b4-de54-9bcd0549f1d4
ms.date: 06/08/2017
ms.openlocfilehash: 2d600012ded93dfdf0a7c6772f990b7d7cc6eb24
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangeisinline-property-publisher"></a>Свойство ShapeRange.IsInline (издатель)

Возвращает константу **MsoTriState** , указывающее, является ли фигура встроенного. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **IsInline**

 переменная _expression_A, представляющий объект **ShapeRange** .


## <a name="example"></a>Пример

В этом примере проверяется первую фигуру (рамке) на первой странице публикации ли встроенного. Если он не установлен, поиск выполняется в рамках фигуры для поиска фигуры, встроенного в пределах рамки. Встроенные фигуры, найденные имеют свойство **ForeColor** , задается красный цвет.


```vb
Dim theShape As Shape 
Dim i As Integer 
 
Set theShape = ActiveDocument.Pages(1).Shapes(1) 
 
If Not theShape.IsInline = True Then 
 With theShape.TextFrame.Story.TextRange 
 If .InlineShapes.Count > 0 Then 
 For i = 1 To .InlineShapes.Count 
 .InlineShapes(i).Select 
 .InlineShapes(i).Fill.ForeColor.RGB = vbRed 
 Next 
 End If 
 End With 
End If
```


