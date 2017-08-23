---
title: "Свойство ParagraphFormat.LeftIndent (издатель)"
keywords: vbapb10.chm5439494
f1_keywords: vbapb10.chm5439494
ms.prod: publisher
api_name: Publisher.ParagraphFormat.LeftIndent
ms.assetid: f9cc3a86-d382-92d7-ec24-d13fc5e3d844
ms.date: 06/08/2017
ms.openlocfilehash: 004d528c43ecb204788846191e1ae165bd75f622
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="paragraphformatleftindent-property-publisher"></a>Свойство ParagraphFormat.LeftIndent (издатель)

Возвращает или задает **Variant** , который представляет значение отступа (в пунктах) для указанного абзацев. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **LeftIndent**

 переменная _expression_A, представляет собой объект- **ParagraphFormat** .


### <a name="return-value"></a>Возвращаемое значение

Variant


## <a name="example"></a>Пример

В этом примере отступы абзаца по позиции курсора 0,5 дюйма. В этом примере предполагается, что курсор находится в текстовом поле.


```vb
Sub IndentParagraph() 
 Selection.TextRange.ParagraphFormat.LeftIndent = 36 
End Sub
```


