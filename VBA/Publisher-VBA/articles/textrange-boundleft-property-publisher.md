---
title: "Свойство TextRange.BoundLeft (издатель)"
keywords: vbapb10.chm5308435
f1_keywords: vbapb10.chm5308435
ms.prod: publisher
api_name: Publisher.TextRange.BoundLeft
ms.assetid: 1ad36906-3dbf-9158-173b-b9047910f6d2
ms.date: 06/08/2017
ms.openlocfilehash: 8391d1ca4bab4dd0e29ff0d37470182cb996cf2f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangeboundleft-property-publisher"></a>Свойство TextRange.BoundLeft (издатель)

Возвращает значение типа **одного** , указывающее, расстояние в пунктах от левого края самые левые страницы по левому краю прямоугольника в диапазоне указанный текст. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **BoundLeft**

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


