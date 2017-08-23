---
title: "Свойство TextRange.BoundTop (издатель)"
keywords: vbapb10.chm5308437
f1_keywords: vbapb10.chm5308437
ms.prod: publisher
api_name: Publisher.TextRange.BoundTop
ms.assetid: f3c2cd42-8d2b-f757-bcbb-140f5e567a1e
ms.date: 06/08/2017
ms.openlocfilehash: 178da1a7cb48c48659f33582da30353f2906e00c
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangeboundtop-property-publisher"></a>Свойство TextRange.BoundTop (издатель)

Возвращает значение типа **одного** , указывающее, расстояние в пунктах от верхнего края верхнего уровня страницы, чтобы верхнего края прямоугольника в диапазоне указанный текст. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **BoundTop**

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


