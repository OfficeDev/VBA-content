---
title: "Свойство ParagraphFormat.LineSpacing (издатель)"
keywords: vbapb10.chm5439504
f1_keywords: vbapb10.chm5439504
ms.prod: publisher
api_name: Publisher.ParagraphFormat.LineSpacing
ms.assetid: cb9abe6a-794c-6a58-2706-e12bbb5a302b
ms.date: 06/08/2017
ms.openlocfilehash: b3cf60154573e6391d343e08dc41da10640e4d95
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="paragraphformatlinespacing-property-publisher"></a>Свойство ParagraphFormat.LineSpacing (издатель)

Возвращает или задает **Variant** , который представляет междустрочным интервалом (в число строк) для указанного абзацев. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **LineSpacing**

 переменная _expression_A, представляет собой объект- **ParagraphFormat** .


### <a name="return-value"></a>Возвращаемое значение

Variant


## <a name="remarks"></a>Заметки

Свойство **[LineSpacingRule](paragraphformat-linespacingrule-property-publisher.md)** задать интервалы строки с определенным значением.


## <a name="example"></a>Пример

В этом примере задается междустрочным интервалом абзаца по позиции курсора в три строки. В этом примере предполагается, что курсор находится в текстовом поле.


```vb
Sub SetLineSpacing() 
 Selection.TextRange.ParagraphFormat.LineSpacing = 3 
End Sub
```


