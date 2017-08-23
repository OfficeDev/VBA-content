---
title: "Свойство ParagraphFormat.LineSpacingRule (издатель)"
keywords: vbapb10.chm5439505
f1_keywords: vbapb10.chm5439505
ms.prod: publisher
api_name: Publisher.ParagraphFormat.LineSpacingRule
ms.assetid: e9855daa-59f4-a4b6-f153-5de515261414
ms.date: 06/08/2017
ms.openlocfilehash: ae0460f04960f8b649928b6f8043070057873ebb
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="paragraphformatlinespacingrule-property-publisher"></a>Свойство ParagraphFormat.LineSpacingRule (издатель)

Возвращает или задает **PbLineSpacingRule** , представляющий междустрочным интервалом для указанного абзацев. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **LineSpacingRule**

 переменная _expression_A, представляет собой объект- **ParagraphFormat** .


### <a name="return-value"></a>Возвращаемое значение

PbLineSpacingRule


## <a name="remarks"></a>Заметки

Значение свойства **LineSpacingRule** может иметь одно из **[PbLineSpacingRule](pblinespacingrule-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

В этом примере форматирование абзаца по позиции курсора двойного интервала.


```vb
Sub SetLineSpacing() 
 Selection.TextRange.ParagraphFormat 
 .LineSpacingRule = pbLineSpacingDouble 
End Sub
```


