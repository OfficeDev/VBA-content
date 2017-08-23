---
title: "Свойство Page.IsTrailing (издатель)"
keywords: vbapb10.chm131101
f1_keywords: vbapb10.chm131101
ms.prod: publisher
api_name: Publisher.Page.IsTrailing
ms.assetid: e0ed15dc-d2e8-d6b7-913d-4e72b2817e88
ms.date: 06/08/2017
ms.openlocfilehash: 378177df9ad09091dceb68d771b1f67f79cd5785
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pageistrailing-property-publisher"></a>Свойство Page.IsTrailing (издатель)

 **Значение true,** Если указанный объект **Page** является конечные страницы из двух страницах. Только для чтения **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **IsTrailing**

 переменная _expression_A, представляющий объект **страницы** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="example"></a>Пример

Следующий пример отображает для каждой страницы ли страница находится в конце строки или основную страницу в публикации.


```vb
Dim objPage As Page 
Dim strPageInfo As String 
For Each objPage In ActiveDocument.Pages 
 strPageInfo = "Page number " &; objPage.PageNumber 
 If objPage.IsLeading Then 
 strPageInfo = strPageInfo &; " is a leading page." &; Chr(13) 
 ElseIf objPage.IsTrailing Then 
 strPageInfo = strPageInfo &; " is a trailing page." &; Chr(13) 
 End If 
 MsgBox strPageInfo 
Next objPage
```


