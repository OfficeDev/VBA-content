---
title: "Свойство ParagraphFormat.RightIndent (издатель)"
keywords: vbapb10.chm5439495
f1_keywords: vbapb10.chm5439495
ms.prod: publisher
api_name: Publisher.ParagraphFormat.RightIndent
ms.assetid: bc3102d3-afc5-3f19-b98a-7f816e374d1a
ms.date: 06/08/2017
ms.openlocfilehash: 227daaea8dc3be420089b17ad3aea47635d7350c
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="paragraphformatrightindent-property-publisher"></a>Свойство ParagraphFormat.RightIndent (издатель)

Возвращает или задает **Variant** , который представляет отступ справа (в пунктах) для указанного абзацев. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **RightIndent**

 переменная _expression_A, представляет собой объект- **ParagraphFormat** .


### <a name="return-value"></a>Возвращаемое значение

Variant


## <a name="example"></a>Пример

В этом примере задается из правого поля отступ справа для всех абзацев в активном документе 2,5. Метод **[InchesToPoints не была назначена](application-inchestopoints-method-publisher.md)** используется для преобразования дюймов в пунктах. В этом примере предполагает наличие по крайней мере один фигуры на первой странице active публикации.


```vb
Sub SetRightIndent() 
 ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.Paragraphs(1).ParagraphFormat _ 
 .RightIndent = InchesToPoints(1) 
End Sub
```


