---
title: "Свойство TextRange.BoundWidth (издатель)"
keywords: vbapb10.chm5308438
f1_keywords: vbapb10.chm5308438
ms.prod: publisher
api_name: Publisher.TextRange.BoundWidth
ms.assetid: bab5053f-958b-9264-9a1e-6f81b5a860b7
ms.date: 06/08/2017
ms.openlocfilehash: 251b94275abd12a406aa98052d161e32c4171810
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangeboundwidth-property-publisher"></a>Свойство TextRange.BoundWidth (издатель)

Возвращает значение типа **одного** , указывающее ширину в пунктах прямоугольника в диапазоне указанный текст. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **BoundWidth**

 переменная _expression_A, представляющий объект **TextRange** .


### <a name="return-value"></a>Возвращаемое значение

Один


## <a name="example"></a>Пример

Следующий пример отображает позицию, ширину и высоту прямоугольника, окружающим текстом в первую фигуру на странице один из активных публикации.


```vb
Dim rngText As TextRange 
Dim strMessage As String 
 
Set rngText = ActiveDocument.Pages(1) _ 
 .Shapes(1).TextFrame.TextRange 
 
With rngText 
 strMessage = "Text frame information" &; vbCrLf _ 
 &; " Distance from left edge of page: " _ 
 &; .BoundLeft &; " points" &; vbCrLf _ 
 &; " Distance from top edge of page: " _ 
 &; .BoundTop &; " points" &; vbCrLf _ 
 &; " Width: " &; .BoundWidth &; " points" &; vbCrLf _ 
 &; " Height: " &; .BoundHeight &; " points" 
End With 
 
MsgBox strMessage
```


