---
title: "Свойство ParagraphFormat.KashidaPercentage (издатель)"
keywords: vbapb10.chm5439513
f1_keywords: vbapb10.chm5439513
ms.prod: publisher
api_name: Publisher.ParagraphFormat.KashidaPercentage
ms.assetid: d62aa512-cce6-2e78-657f-51ff1b2cbcf8
ms.date: 06/08/2017
ms.openlocfilehash: 3abd4fa867e2cc2be1129909b9dc0cba3ddd7362
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="paragraphformatkashidapercentage-property-publisher"></a>Свойство ParagraphFormat.KashidaPercentage (издатель)

Возвращает или задает типа **Long** , указывающее процент, с помощью которого кашиды должны быть в тексте для указанного абзацев. Допустимые значения: от 0 до 100. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **KashidaPercentage**

 переменная _expression_A, представляет собой объект- **ParagraphFormat** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="remarks"></a>Заметки

Свойство **[Alignment](paragraphformat-alignment-property-publisher.md)** указанного абзацев должен иметь значение **pbParagraphAlignmentKashida** или **KashidaPercentage** свойство будет пропущено.


## <a name="example"></a>Пример

В следующем примере задается абзацы в форму одно на странице один из активных публикации для выравнивания кашида и указывает, что кашиды должны быть в тексте на 20 процентов.


```vb
With ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.ParagraphFormat 
 .Alignment = pbParagraphAlignmentKashida 
 .KashidaPercentage = 20 
End With
```


