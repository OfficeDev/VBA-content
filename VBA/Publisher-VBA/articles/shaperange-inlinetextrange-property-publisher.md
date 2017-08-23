---
title: "Свойство ShapeRange.InlineTextRange (издатель)"
keywords: vbapb10.chm2294023
f1_keywords: vbapb10.chm2294023
ms.prod: publisher
api_name: Publisher.ShapeRange.InlineTextRange
ms.assetid: 5d7f3dfa-3e23-85c6-50cf-a6f960ccabfc
ms.date: 06/08/2017
ms.openlocfilehash: 9eca5c01bc7c2363cda34d67b6ab991cd5f34b97
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangeinlinetextrange-property-publisher"></a>Свойство ShapeRange.InlineTextRange (издатель)

Возвращает объект **[TextRange](textrange-object-publisher.md)** , показывающая положение фигуры встроенного в его содержащего диапазон текста. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **InlineTextRange**

 переменная _expression_A, представляющий объект **ShapeRange** .


## <a name="remarks"></a>Заметки

Диапазон возвращаемый текст будет содержать один объект, представляющий встроенная фигура. Если фигура не является встроенной, возвращается ошибка автоматизации.


## <a name="example"></a>Пример

В следующем примере выполняется поиск первую фигуру (текстовое поле) на первой странице публикации и определяет, содержит ли диапазон текста в текстовом поле встроенных фигур. Если обнаружены встроенных фигур, свойство **InlineTextRange** используется для представления встроенная фигура после вставки блока текста.


```vb
Dim theShape As Shape 
Dim theTextRange As TextRange 
Dim i As Integer 
 
Set theShape = ActiveDocument.Pages(1).Shapes(1) 
 
If Not theShape.IsInline = True Then 
 With theShape.TextFrame.Story.TextRange 
 If .InlineShapes.Count > 0 Then 
 Set theTextRange = theShape.TextFrame.Story.TextRange 
 For i = 1 To .InlineShapes.Count 
 With .InlineShapes(i) 
 .InlineTextRange.InsertAfter (" (Figure " &; i &; ") ") 
 End With 
 Next 
 End If 
 End With 
End If
```


