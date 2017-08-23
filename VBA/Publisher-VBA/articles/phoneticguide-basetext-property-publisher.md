---
title: "Свойство PhoneticGuide.BaseText (издатель)"
keywords: vbapb10.chm6160391
f1_keywords: vbapb10.chm6160391
ms.prod: publisher
api_name: Publisher.PhoneticGuide.BaseText
ms.assetid: e59ef54f-c650-1a3e-717b-b4b603f312c1
ms.date: 06/08/2017
ms.openlocfilehash: dc33a6c3567bb72b2ebb284c065edf725fbefc81
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="phoneticguidebasetext-property-publisher"></a>Свойство PhoneticGuide.BaseText (издатель)

Возвращает **строку** , представляющую текст, к которому относится указанный текст фонетическое. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **BaseText**

 переменная _expression_A, представляет собой объект- **PhoneticGuide** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="example"></a>Пример

В этом примере добавляется фонетическое текст для выделения и отображается текст, к которому применяется фонетическое текста, содержит текст изначально выбранные. В этом примере предполагается, что выделенный текст. Если текст не выделен, в окне сообщения будет пустым.


```vb
Sub AddPhoneticText() 
 With Selection.TextRange.Fields.AddPhoneticGuide _ 
 (Range:=Selection.TextRange, Text:="tray sheek") 
 MsgBox "The base text is " &; .PhoneticGuide.BaseText 
 End With 
End Sub
```


