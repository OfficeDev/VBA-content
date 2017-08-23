---
title: "Свойство ParagraphFormat.SpaceAfter (издатель)"
keywords: vbapb10.chm5439496
f1_keywords: vbapb10.chm5439496
ms.prod: publisher
api_name: Publisher.ParagraphFormat.SpaceAfter
ms.assetid: 52f65636-862d-442e-e66f-5ff5c79ee7b0
ms.date: 06/08/2017
ms.openlocfilehash: 5e0b2d9cdd1ffc5eaaad213a133425148068ee9a
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="paragraphformatspaceafter-property-publisher"></a>Свойство ParagraphFormat.SpaceAfter (издатель)

Возвращает или задает **Variant** , который представляет значение интервала (в пунктах) после одного или нескольких абзацев. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **SpaceAfter**

 переменная _expression_A, представляет собой объект- **ParagraphFormat** .


### <a name="return-value"></a>Возвращаемое значение

Variant


## <a name="example"></a>Пример

В этом примере задается интервал до и после третий абзац в первую фигуру на первой странице active публикации на 6 пунктов.


```vb
Sub SetSpacingBeforeAfterParagraph() 
 With ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.Paragraphs(3).ParagraphFormat 
 .SpaceBefore = 6 
 .SpaceAfter = 6 
 End With 
End Sub
```

В этом примере задается интервал до и после все абзацы в первую фигуру на первой странице active публикации на 6 пунктов.




```vb
Sub SetSpacingBeforeAfterAllParagraph() 
 With ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.ParagraphFormat 
 .SpaceBefore = 12 
 .SpaceAfter = 6 
 End With 
End Sub
```


