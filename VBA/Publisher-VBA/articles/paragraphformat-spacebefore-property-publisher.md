---
title: "Свойство ParagraphFormat.SpaceBefore (издатель)"
keywords: vbapb10.chm5439497
f1_keywords: vbapb10.chm5439497
ms.prod: publisher
api_name: Publisher.ParagraphFormat.SpaceBefore
ms.assetid: ed19a927-67e4-a1b3-06f8-1035c4b0815a
ms.date: 06/08/2017
ms.openlocfilehash: e16f44f2071c64ecde25912f6dc45c96c9d54e09
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="paragraphformatspacebefore-property-publisher"></a>Свойство ParagraphFormat.SpaceBefore (издатель)

Возвращает или задает **Variant** , который представляет значение интервала (в пунктах) до одного или нескольких абзацев. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **SpaceBefore**

 переменная _expression_A, представляет собой объект- **ParagraphFormat** .


### <a name="return-value"></a>Возвращаемое значение

Variant


## <a name="example"></a>Пример

В этом примере задается интервал до и после третий абзац в первую фигуру на первой странице active публикации на 6 пунктов. В этом примере предполагается, что имеется по крайней мере один фигуры на первой странице active публикации.


```vb
Sub SetSpacingBeforeAfterParagraph() 
 With ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.Paragraphs(3).ParagraphFormat 
 .SpaceBefore = 6 
 .SpaceAfter = 6 
 End With 
End Sub
```

В этом примере задается интервал до и после все абзацы в первую фигуру на первой странице active публикации на 6 пунктов. В этом примере предполагается, что имеется по крайней мере один фигуры на первой странице active публикации.




```vb
Sub SetSpacingBeforeAfterAllParagraph() 
 With ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.ParagraphFormat 
 .SpaceBefore = 12 
 .SpaceAfter = 6 
 End With 
End Sub
```


