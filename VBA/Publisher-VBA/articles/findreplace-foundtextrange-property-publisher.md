---
title: "Свойство FindReplace.FoundTextRange (издатель)"
keywords: vbapb10.chm8323075
f1_keywords: vbapb10.chm8323075
ms.prod: publisher
api_name: Publisher.FindReplace.FoundTextRange
ms.assetid: 8d0d3177-2d32-7df6-8b88-b354ec0a3d7b
ms.date: 06/08/2017
ms.openlocfilehash: a5a9552f4bc53681f0468648c5f0524e56f59fe8
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="findreplacefoundtextrange-property-publisher"></a>Свойство FindReplace.FoundTextRange (издатель)

Возвращает объект **TextRange** , который представляет найденный текст или замены текста операции поиска. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **FoundTextRange**

 переменная _expression_A, представляет собой объект- **FindReplace** .


### <a name="return-value"></a>Возвращаемое значение

TextRange


## <a name="remarks"></a>Заметки

Фактический объект **TextRange** , возвращенный свойством **FoundTextRange** определяется значение свойства **ReplaceScope** . В следующей таблице перечислены соответствующие значения этих свойств.



| для **ReplaceScope** = **pbReplaceScopeAll**| **FoundTextRange** = Empty | | для **ReplaceScope** = **pbReplaceScopeNone**| **FoundTextRange** = Find диапазон текста | | для **ReplaceScope** = **pbReplaceScopeOne**| **FoundTextRange** = диапазон текста заменить | Если **ReplaceScope** **pbReplaceScopeAll**, свойство **FoundTextRange** будет пустым. Любая попытка обратиться к ней возвращает «Отказано в доступе.» Способ работы с диапазон текста, поиск текста — задайте свойству **ReplaceScope** значение **pbReplaceScopeNone** или **pbReplaceScopeOne** и получать доступ к диапазон текста поиска или замененный текст для каждого повтора найден.


## <a name="example"></a>Пример

Если **ReplaceScope** **pbReplaceScopeNone**, **FoundTextRange** возвращает диапазон текста, поиск текста. В следующем примере показано, как может осуществляться атрибуты шрифта текстового диапазона поиска при **ReplaceScope** задано значение **pbReplaceScopeNone**.


```vb
With TextRange.Find 
 .Clear 
 .FindText = "important" 
 .ReplaceScope = pbReplaceScopeNone 
 Do While .Execute = True 
 'The FoundTextRange contains the word "important". 
 If .FoundTextRange.Font.Italic = msoFalse Then 
 .FoundTextRange.Font.Italic = msoTrue 
 End If 
 Loop 
End With
```

Если **ReplaceScope** **pbReplaceScopeOne**, диапазон текста, поиск текста будет заменен. Таким образом свойство **FoundTextRange** возвращает диапазон текста замещающий текст. В следующем примере показано, как может осуществляться атрибуты шрифта диапазона текст заменен при **ReplaceScope** задано значение **pbReplaceScopeOne**. 




```vb
With Document.Find 
 .Clear 
 .FindText = "important" 
 .ReplaceWithText = "urgent" 
 .ReplaceScope = pbReplaceScopeOne 
 Do While .Execute = True 
 'The FoundTextRange contains the word "urgent". 
 If .FoundTextRange.Font.Bold = msoFalse Then 
 .FoundTextRange.Font.Bold = msoTrue 
 End If 
 Loop 
End With
```

В этом примере заменяет каждый пример слово «довольно странно» со словом «разрешение» и применяет курсивом и полужирным шрифтом, замененный текст. 




```vb
Dim objDocument As Document 
 
Set objDocument = ActiveDocument 
With objDocument.Find 
 .Clear 
 .FindText = "bizarre" 
 .ReplaceWithText = "strange" 
 .ReplaceScope = pbReplaceScopeOne 
 Do While .Execute = True 
 .FoundTextRange.Font.Italic = msoTrue 
 .FoundTextRange.Font.Bold = msoTrue 
 Loop 
End With
```

В этом примере находит все вхождения слова «важно» и курсивом его.




```vb
Dim objTextRange As TextRange 
 
Set objTextRange = ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange 
With objTextRange.Find 
 .Clear 
 .FindText = "important" 
 .ReplaceScope = pbReplaceScopeNone 
 Do While .Execute = True 
 .FoundTextRange.Font.Italic = msoTrue 
 Loop 
End With
```


