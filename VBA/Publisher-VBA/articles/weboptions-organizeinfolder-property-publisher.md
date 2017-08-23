---
title: "Свойство WebOptions.OrganizeInFolder (издатель)"
keywords: vbapb10.chm8257542
f1_keywords: vbapb10.chm8257542
ms.prod: publisher
api_name: Publisher.WebOptions.OrganizeInFolder
ms.assetid: f09ac701-d8d8-a58f-965c-bd5e4b69820c
ms.date: 06/08/2017
ms.openlocfilehash: a7b4048ae24201d4a306b26c85db84661cc10241
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="weboptionsorganizeinfolder-property-publisher"></a>Свойство WebOptions.OrganizeInFolder (издатель)

Возвращает или задает **логическое** значение, указывающее, является ли веб-публикации будут сохраняться плоской структуры или иерархическую структуру. Если **значение False**, все файлы в веб-публикации будет сохранен в виде плоской структуры в корневой папке. Если **значение True**, файлы будут сохранены в иерархической структуры в корневой папке. Значение по умолчанию — **True**. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **OrganizeInFolder**

 переменная _expression_A, представляющий объект **WebOptions** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="example"></a>Пример

Следующий пример указывает, сохранить все файлы в веб-публикации в плоской структуры в корневой папке.


```vb
Dim theWO As WebOptions 
 
Set theWO = Application.WebOptions 
 
With theWO 
 .OrganizeInFolder = False 
End With
```


