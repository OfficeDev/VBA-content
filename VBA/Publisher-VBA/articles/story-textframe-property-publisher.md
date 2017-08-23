---
title: "Свойство Story.TextFrame (издатель)"
keywords: vbapb10.chm5832709
f1_keywords: vbapb10.chm5832709
ms.prod: publisher
api_name: Publisher.Story.TextFrame
ms.assetid: bb6ce510-068c-27c2-9df0-a709ab46db2e
ms.date: 06/08/2017
ms.openlocfilehash: 80f5a2fe27a346c63fe2d4ba46e6debfc9f06286
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="storytextframe-property-publisher"></a>Свойство Story.TextFrame (издатель)

Возвращает объект **[TextFrame](textframe-object-publisher.md)** , представляющий текст в фигуру и свойства, которые управляют поля и ориентацию текста.


## <a name="syntax"></a>Синтаксис

 _выражение_. **TextFrame**

 переменная _expression_A, представляет собой объект- **материала** .


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


