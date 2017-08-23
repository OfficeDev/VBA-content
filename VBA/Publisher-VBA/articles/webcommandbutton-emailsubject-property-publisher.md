---
title: "Свойство WebCommandButton.EmailSubject (издатель)"
keywords: vbapb10.chm3932168
f1_keywords: vbapb10.chm3932168
ms.prod: publisher
api_name: Publisher.WebCommandButton.EmailSubject
ms.assetid: 4d29dacd-0da6-c706-515e-219daf5e349d
ms.date: 06/08/2017
ms.openlocfilehash: 6fecafe7fc6c9a348232eee1e61090ce41711575
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webcommandbuttonemailsubject-property-publisher"></a>Свойство WebCommandButton.EmailSubject (издатель)

Задает или возвращает **строку** , представляющую тему сообщения электронной почты, созданные для обработки данных веб-форм. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **EmailSubject**

 переменная _expression_A, представляет собой объект- **WebCommandButton** .


## <a name="example"></a>Пример

В этом примере Publisher для обработки данных на веб-форму в текущей публикации, отправив сообщение электронной почты с темой на указанный адрес электронной почты.


```vb
Sub WebFormData() 
 With ThisDocument.Pages(1).Shapes(1).WebCommandButton 
 .DataRetrievalMethod = pbSubmitDataRetrievalEmail 
 .EmailAddress = "someone@example.com" 
 .EmailSubject = "Web form data" 
 End With 
End Sub
```


