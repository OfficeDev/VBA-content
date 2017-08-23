---
title: "Свойство TextRange.MajorityParagraphFormat (издатель)"
keywords: vbapb10.chm5308468
f1_keywords: vbapb10.chm5308468
ms.prod: publisher
api_name: Publisher.TextRange.MajorityParagraphFormat
ms.assetid: d67e81fe-ab9b-8bfd-c31d-76feb1b6e15b
ms.date: 06/08/2017
ms.openlocfilehash: dbc1b44a9157a8862e38b7a78829cb5ae2b72876
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangemajorityparagraphformat-property-publisher"></a>Свойство TextRange.MajorityParagraphFormat (издатель)

Возвращает объект **[ParagraphFormat](paragraphformat-object-publisher.md)** , представляющий форматирование абзаца, применяемые к большая часть абзацев в диапазон текста.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MajorityParagraphFormat**

 переменная _expression_A, представляющий объект **TextRange** .


### <a name="return-value"></a>Возвращаемое значение

ParagraphFormat


## <a name="example"></a>Пример

В этом примере применяется форматирование абзаца, применяемые к большинство абзацев в первую фигуру для абзацев в эту фигуру на первой странице активного документа. В этом примере предполагается, что на странице один из активных публикации имеются по крайней мере двух фигур.


```vb
Sub SetFontName() 
 Dim fmt As ParagraphFormat 
 With ActiveDocument.Pages(1) 
 Set fmt = .Shapes(1).TextFrame.TextRange _ 
 .MajorityParagraphFormat 
 .Shapes(2).TextFrame.TextRange.ParagraphFormat = fmt 
 End With 
End Sub
```


