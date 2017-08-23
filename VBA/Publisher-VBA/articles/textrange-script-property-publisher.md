---
title: "Свойство TextRange.Script (издатель)"
keywords: vbapb10.chm5308484
f1_keywords: vbapb10.chm5308484
ms.prod: publisher
api_name: Publisher.TextRange.Script
ms.assetid: 54e5a19f-9cb0-0fbc-5ebe-cd4db6c0de8e
ms.date: 06/08/2017
ms.openlocfilehash: 5bb2f47b944529b67e7bc3dcfbfbb5acd30a5517
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangescript-property-publisher"></a>Свойство TextRange.Script (издатель)

Возвращает константу **PbFontScriptType** , представляющий начертания шрифта для диапазона текста. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Сценарий**

 переменная _expression_A, представляющий объект **TextRange** .


### <a name="return-value"></a>Возвращаемое значение

PbFontScriptType


## <a name="remarks"></a>Заметки

Значение свойства **скрипт** может иметь одно из **[PbFontScriptType](pbfontscripttype-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

В этом примере отображается сообщение при ASCII Латинская сценария шрифта, используемого в диапазоне указанный текст. В этом примере предполагает наличие по крайней мере один фигуры на первой странице active публикации.


```vb
Sub DisplayScriptType() 
 If ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange _ 
 .Script = pbFontScriptAsciiLatin Then 
 MsgBox "The font script you are using is ASCII Latin." 
 End If 
End Sub
```


