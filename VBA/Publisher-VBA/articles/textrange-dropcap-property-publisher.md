---
title: "Свойство TextRange.DropCap (издатель)"
keywords: vbapb10.chm5308472
f1_keywords: vbapb10.chm5308472
ms.prod: publisher
api_name: Publisher.TextRange.DropCap
ms.assetid: a5c29dd4-62f4-39fb-4b76-390d62bd8e32
ms.date: 06/08/2017
ms.openlocfilehash: 4d4acfde3b91a1daf6515bf1c739b9526e4771bc
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangedropcap-property-publisher"></a>Свойство TextRange.DropCap (издатель)

Возвращает объект **[буквицу](dropcap-object-publisher.md)** , представляющий буквицы для абзацев в элементе frame указанный текст.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Буквицу**

 переменная _expression_A, представляющий объект **TextRange** .


### <a name="return-value"></a>Возвращаемое значение

Буквицу


## <a name="example"></a>Пример

В этом примере применяется настраиваемых добавленном капитала, три строки высокой и занимает первые три символа все абзацы в элементе frame указанный текст.


```vb
Sub SetDropCap() 
 With ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange 
 .DropCap.ApplyCustomDropCap FontName:="Snap ITC", _ 
 Bold:=True, Size:=3, Span:=3 
 With .ParagraphFormat 
 .SpaceBefore = 6 
 .SpaceAfter = 6 
 End With 
 End With 
End Sub
```


