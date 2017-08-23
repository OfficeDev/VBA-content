---
title: "Свойство WebCommandButton.EmailAddress (издатель)"
keywords: vbapb10.chm3932167
f1_keywords: vbapb10.chm3932167
ms.prod: publisher
api_name: Publisher.WebCommandButton.EmailAddress
ms.assetid: 8961e459-1ce1-558a-2450-c3b8da2d5559
ms.date: 06/08/2017
ms.openlocfilehash: 1d425dfc3ddee62bae6d8d790d1deae84eae583b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webcommandbuttonemailaddress-property-publisher"></a>Свойство WebCommandButton.EmailAddress (издатель)

Задает или возвращает **строку** представляющее адрес электронной почты для использования при обработке данных веб-форм. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **EmailAddress**

 переменная _expression_A, представляющий объект **WebCommandButton** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="example"></a>Пример

В этом примере Microsoft Publisher для обработки данных на веб-форму в текущей публикации, отправив сообщение электронной почты на указанный адрес электронной почты.


```vb
Sub WebFormData() 
 With ThisDocument.Pages(1).Shapes(1).WebCommandButton 
 .DataRetrievalMethod = pbSubmitDataRetrievalEmail 
 .EmailAddress = "someone@example.com" 
 .EmailSubject = "Web form data" 
 End With 
End Sub
```


