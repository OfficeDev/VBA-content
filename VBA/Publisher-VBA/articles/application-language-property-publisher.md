---
title: "Свойство Application.Language (издатель)"
keywords: vbapb10.chm131091
f1_keywords: vbapb10.chm131091
ms.prod: publisher
api_name: Publisher.Application.Language
ms.assetid: 2fcfbec9-0c84-43d5-8c53-5b73bca17e3d
ms.date: 06/08/2017
ms.openlocfilehash: 2ed5f2535a9ceeb5671c5ac3e49449ecb7f3d16b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationlanguage-property-publisher"></a>Свойство Application.Language (издатель)

Возвращает значение типа **Long** , представляющее язык, выбранный для пользовательского интерфейса Microsoft Publisher. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Язык**

 переменная _expression_A, представляющий объект **приложения** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="remarks"></a>Заметки

Значение свойства **LanguageID** может иметь одно из ** [MsoLanguageID](http://msdn.microsoft.com/library/65ea40f0-9a09-3d76-1519-4acddcc5f367%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.


## <a name="example"></a>Пример

В этом примере отображается сообщение о том, является ли язык, выбранный для пользовательского интерфейса Publisher английский (США).


```vb
Sub LangSetting() 
 If Application.Language = msoLanguageIDEnglishUS Then 
 MsgBox "The user interface language is U.S. English." 
 Else 
 MsgBox "The user interface language is not U.S. English." 
 End If 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

