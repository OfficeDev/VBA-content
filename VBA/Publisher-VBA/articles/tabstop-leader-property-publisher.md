---
title: "Свойство TabStop.Leader (издатель)"
keywords: vbapb10.chm5636101
f1_keywords: vbapb10.chm5636101
ms.prod: publisher
api_name: Publisher.TabStop.Leader
ms.assetid: a788bdc8-8ab3-fcd3-931a-a5b83db93991
ms.date: 06/08/2017
ms.openlocfilehash: 910a2e2314d059fa64f99debcb2579dfd2c63fe3
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="tabstopleader-property-publisher"></a>Свойство TabStop.Leader (издатель)

Задает или возвращает константу **PbTabLeaderType** , который представляет заполнитель для позиции табуляции. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Ведущий сотрудник**

 переменная _expression_A, представляет собой объект- **TabStop** .


### <a name="return-value"></a>Возвращаемое значение

PbTabLeaderType


## <a name="remarks"></a>Заметки

Значение свойства **ведущий** может иметь одно из **[PbTabLeaderType](pbtableadertype-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

В этом примере изменяется символ табуляции ведущий выбранного абзацев тире. В этом примере предполагается, что выбранный абзац содержит по крайней мере один позиции табуляции.


```vb
Sub SetLeaderTab() 
 Selection.TextRange.ParagraphFormat _ 
 .Tabs(1).Leader = pbTabLeaderDashes 
End Sub
```

В этом примере изменяется символ табуляции ведущий первого абзаца в диапазоне указанный текст подчеркивание. В этом примере предполагается, что указанный абзац содержит по крайней мере один позиции табуляции.




```vb
Sub SetNewTabLeader() 
 ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange.Paragraphs(1) _ 
 .ParagraphFormat.Tabs(1).Leader = pbTabLeaderLine 
End Sub
```


