---
title: "Свойство WebCommandButton.ActionURL (издатель)"
keywords: vbapb10.chm3932163
f1_keywords: vbapb10.chm3932163
ms.prod: publisher
api_name: Publisher.WebCommandButton.ActionURL
ms.assetid: ede9b18f-1be1-9572-9b78-7dbe0817cfe7
ms.date: 06/08/2017
ms.openlocfilehash: 6dddc9d5b1bc6580aa9942cd203204f22b5fbe67
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webcommandbuttonactionurl-property-publisher"></a>Свойство WebCommandButton.ActionURL (издатель)

Возвращает или задает **строку** , представляющую URL-адреса сценариев на стороне сервера, выполняются в ответ на нажатие кнопки Отправить. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ActionURL**

 переменная _expression_A, представляет собой объект- **WebCommandButton** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="remarks"></a>Заметки

Значение по умолчанию для свойства **ActionURL** : «http://example.microsoft.com/~user/ispscript.cgi». Это свойство игнорируется для сброса кнопок.


## <a name="example"></a>Пример

В этом примере создается кнопки Отправить форму Web и задает путь и имя скрипта для запуска при нажатии кнопки.


```vb
Sub CreateActionWebButton() 
 With ActiveDocument.Pages(1).Shapes.AddWebControl _ 
 (Type:=pbWebControlCommandButton, Left:=150, _ 
 Top:=150, Width:=75, Height:=36).WebCommandButton 
 .ButtonText = "Submit" 
 .ButtonType = pbCommandButtonSubmit 
 .ActionURL = "http://www.tailspintoys.com/" &; _ 
 "scripts/ispscript.cgi" 
 End With 
End Sub
```


