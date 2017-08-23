---
title: "Свойство Hyperlink.EmailSubject (издатель)"
keywords: vbapb10.chm4587524
f1_keywords: vbapb10.chm4587524
ms.prod: publisher
api_name: Publisher.Hyperlink.EmailSubject
ms.assetid: 16b60648-56fe-b8ba-3424-0dd6e88727e6
ms.date: 06/08/2017
ms.openlocfilehash: 522ff5131b554582b8a2030e95d4590057f1a6c5
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="hyperlinkemailsubject-property-publisher"></a>Свойство Hyperlink.EmailSubject (издатель)

Задает или возвращает **строку** , представляющую тему сообщения электронной почты, созданные для обработки данных веб-форм. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **EmailSubject**

 переменная _expression_A, представляющий объект **гиперссылки** .


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


