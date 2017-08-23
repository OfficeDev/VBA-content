---
title: "Свойство TextRange.Length (издатель)"
keywords: vbapb10.chm5308432
f1_keywords: vbapb10.chm5308432
ms.prod: publisher
api_name: Publisher.TextRange.Length
ms.assetid: 003b4ad1-2c09-17c9-279b-b1cf2ebdb40a
ms.date: 06/08/2017
ms.openlocfilehash: a1e15f47df0253406d80cddbcc0d81ab4a8d9976
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangelength-property-publisher"></a>Свойство TextRange.Length (издатель)

Возвращает значение типа **Long** , указывающее длину диапазона указанный текст в символах. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Длина**

 переменная _expression_A, представляющий объект **TextRange** .


## <a name="example"></a>Пример

В этом примере задает размер шрифта фрагмент текста на странице двух до 48 точек Если надпись содержит более пяти символов, или задает размер шрифта 72 точки, если рамки содержит более пяти символов.


```vb
With ActiveDocument.Pages(2).Shapes(1) _ 
 .TextFrame.TextRange 
 If .Length > 5 Then 
 .Font.Size = 48 
 Else 
 .Font.Size = 72 
 End If 
End With
```


