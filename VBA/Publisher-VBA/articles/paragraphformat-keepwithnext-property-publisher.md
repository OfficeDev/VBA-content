---
title: "Свойство ParagraphFormat.KeepWithNext (издатель)"
keywords: vbapb10.chm5439538
f1_keywords: vbapb10.chm5439538
ms.prod: publisher
api_name: Publisher.ParagraphFormat.KeepWithNext
ms.assetid: fb49169d-4718-8ee6-6468-b7cbc8b8a774
ms.date: 06/08/2017
ms.openlocfilehash: ca74f092fb6e76a6716f76bd611de974760fc430
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="paragraphformatkeepwithnext-property-publisher"></a>Свойство ParagraphFormat.KeepWithNext (издатель)

Задает или возвращает **MsoTriState** , которое указывает, остается ли следующий абзац в текстовом поле же указанного абзаца. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **KeepWithNext**

 переменная _expression_A, представляет собой объект- **ParagraphFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Keep с Далее предназначен для предотвращения длина заголовков в документе. Чтобы сделать, который может этому свойству присвоено значение **msoTrue** для всех заголовков.

Значение по умолчанию для этого свойства — **msoFalse**.


## <a name="example"></a>Пример

В этом примере задается свойство **KeepWithNext** **msoTrue** для указанного объекта **ParagraphFormat** .


```vb
Dim objParaForm As ParagraphFormat 
Set objParaForm = ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Paragraphs(1).ParagraphFormat 
objParaForm.KeepWithNext = msoTrue
```


