---
title: "Свойство WebPageOptions.Keywords (издатель)"
keywords: vbapb10.chm544772
f1_keywords: vbapb10.chm544772
ms.prod: publisher
api_name: Publisher.WebPageOptions.Keywords
ms.assetid: 8dd7b073-747e-a6f6-a20d-0b3e3d9a27b8
ms.date: 06/08/2017
ms.openlocfilehash: 3513d3d619dbee19e1477f53206bd2454a534538
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webpageoptionskeywords-property-publisher"></a>Свойство WebPageOptions.Keywords (издатель)

Возвращает или задает **строку** , представляющую ключевые слова для веб-страницы в веб-публикации. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Ключевые слова**

 переменная _expression_A, представляет собой объект- **WebPageOptions** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="example"></a>Пример

В следующем примере задается ключевые слова для страницы четыре active публикации.


```vb
Dim theWPO As WebPageOptions 
 
Set theWPO = ActiveDocument.Pages(4).WebPageOptions 
 
With theWPO 
 .Keywords = "software, hardware, computers" 
End With
```


