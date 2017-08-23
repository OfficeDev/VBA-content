---
title: "Свойство Page.WebPageOptions (издатель)"
keywords: vbapb10.chm393264
f1_keywords: vbapb10.chm393264
ms.prod: publisher
api_name: Publisher.Page.WebPageOptions
ms.assetid: c2e3ee01-5b49-e83c-a68b-a4d526da0215
ms.date: 06/08/2017
ms.openlocfilehash: 3467f932b470b4028cd9695521b4f8115dc0024e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagewebpageoptions-property-publisher"></a>Свойство Page.WebPageOptions (издатель)

Возвращает объект **[WebPageOptions](webpageoptions-object-publisher.md)** , который представляет свойства одного веб-страницы в веб-публикации. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **WebPageOptions**

 переменная _expression_A, представляющий объект **Page** .


### <a name="return-value"></a>Возвращаемое значение

WebPageOptions


## <a name="example"></a>Пример

В следующем примере задается описание и звуковое сопровождение для четвертой странице active веб-публикации.


```vb
With ActiveDocument.Pages(4).WebPageOptions 
 .Description = "Company Profile" 
 .BackgroundSound = "C:\CompanySounds\corporate_jingle.wav" 
End With 

```


