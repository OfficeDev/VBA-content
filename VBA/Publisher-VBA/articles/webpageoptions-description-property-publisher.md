---
title: "Свойство WebPageOptions.Description (издатель)"
keywords: vbapb10.chm544771
f1_keywords: vbapb10.chm544771
ms.prod: publisher
api_name: Publisher.WebPageOptions.Description
ms.assetid: dfd18427-c70d-7232-191e-a6332a89c3fe
ms.date: 06/08/2017
ms.openlocfilehash: 00ef25e1116f7f5c1cbbf46412aeb9cedb290874
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webpageoptionsdescription-property-publisher"></a>Свойство WebPageOptions.Description (издатель)

Возвращает или задает **строку** , представляющую описание веб-страницы в веб-публикации. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Описание**

 переменная _expression_A, представляет собой объект- **WebPageOptions** .


## <a name="example"></a>Пример

В этом примере задается описание для страницы, два активных веб-публикации.


```vb
Dim theWPO As WebPageOptions 
 
Set theWPO = ActiveDocument.Pages(2).WebPageOptions 
 
With theWPO 
 .Description = "Company Profile" 
End With
```


