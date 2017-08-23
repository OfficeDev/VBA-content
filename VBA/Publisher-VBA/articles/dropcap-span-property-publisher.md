---
title: "Свойство DropCap.Span (издатель)"
keywords: vbapb10.chm5505033
f1_keywords: vbapb10.chm5505033
ms.prod: publisher
api_name: Publisher.DropCap.Span
ms.assetid: 00c51e48-5bbc-13e9-2d0c-e8993f753bbe
ms.date: 06/08/2017
ms.openlocfilehash: 2ed41c2aefdf6e34669c8c6680a6ec0475c58be7
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="dropcapspan-property-publisher"></a>Свойство DropCap.Span (издатель)

Возвращает или задает **Long** , представляющее номер букв, включенных в указанный буквицы. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Диапазон**

 переменная _expression_A, представляет собой объект- **буквицу** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="example"></a>Пример

В этом примере создается пользовательский буквицы, пять строк, занимает первые три символа абзацев в диапазон текста и возникает один линия над первой строки.


```vb
Sub SetDropCapSpan() 
 With ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.DropCap 
 .Size = 5 
 .Span = 3 
 .LinesUp = 1 
 End With 
End Sub
```


