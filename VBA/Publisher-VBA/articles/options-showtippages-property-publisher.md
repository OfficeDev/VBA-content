---
title: "Свойство Options.ShowTipPages (издатель)"
keywords: vbapb10.chm1048609
f1_keywords: vbapb10.chm1048609
ms.prod: publisher
api_name: Publisher.Options.ShowTipPages
ms.assetid: 44f91cf1-68e3-0755-3114-5dc41a2e4eba
ms.date: 06/08/2017
ms.openlocfilehash: 2ba06258820910ab7b19f57d37f7620fac599ea6
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="optionsshowtippages-property-publisher"></a>Свойство Options.ShowTipPages (издатель)

 **Значение true** для Microsoft Publisher для отображения страниц советов в выносках. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ShowTipPages**

 переменная _expression_A, представляет собой объект- **Параметры** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="example"></a>Пример

В этом примере отключается отображение страниц советов в выносках.


```vb
Sub DontShowTipPages() 
 Options.ShowTipPages = False 
End Sub
```


