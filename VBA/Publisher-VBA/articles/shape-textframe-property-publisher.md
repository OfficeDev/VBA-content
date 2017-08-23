---
title: "Свойство Shape.TextFrame (издатель)"
keywords: vbapb10.chm2228304
f1_keywords: vbapb10.chm2228304
ms.prod: publisher
api_name: Publisher.Shape.TextFrame
ms.assetid: fc654905-d56b-9a6c-28fa-4b54bf2a8686
ms.date: 06/08/2017
ms.openlocfilehash: 059476be6308123f1942b2330b9db8afeaa344a7
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapetextframe-property-publisher"></a>Свойство Shape.TextFrame (издатель)

Возвращает объект **[TextFrame](textframe-object-publisher.md)** , представляющий текст в фигуру и свойства, которые управляют поля и ориентацию текста.


## <a name="syntax"></a>Синтаксис

 _выражение_. **TextFrame**

 переменная _expression_A, представляющий объект **фигуры** .


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


