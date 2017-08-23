---
title: "Свойство Shape.IsInline (издатель)"
keywords: vbapb10.chm5308692
f1_keywords: vbapb10.chm5308692
ms.prod: publisher
api_name: Publisher.Shape.IsInline
ms.assetid: 5c5c6181-070f-2a66-8d70-2d6372cb365e
ms.date: 06/08/2017
ms.openlocfilehash: b081e9ff9fdd307454bff7f2ef7e9326826e673e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapeisinline-property-publisher"></a>Свойство Shape.IsInline (издатель)

Возвращает константу **MsoTriState** , указывающее, является ли фигура встроенного (содержащиеся в прогона текста). Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **IsInline**

 переменная _expression_A, представляющий объект **фигуры** .


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


