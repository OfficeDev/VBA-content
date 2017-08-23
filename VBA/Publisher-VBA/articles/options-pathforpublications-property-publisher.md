---
title: "Свойство Options.PathForPublications (издатель)"
keywords: vbapb10.chm1048597
f1_keywords: vbapb10.chm1048597
ms.prod: publisher
api_name: Publisher.Options.PathForPublications
ms.assetid: d33d5eab-eb52-b533-8968-31ddb5e12d99
ms.date: 06/08/2017
ms.openlocfilehash: b6176589f08ce43a16b1615459d2338a44df2e41
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="optionspathforpublications-property-publisher"></a>Свойство Options.PathForPublications (издатель)

Возвращает **строку** , представляющую папки по умолчанию для публикаций. Чтение.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PathForPublications**

 переменная _expression_A, представляет собой объект- **Параметры** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="example"></a>Пример

В этом примере возвращается текущий путь по умолчанию для публикаций (соответствует путь по умолчанию на вкладке **Общие** в диалоговом окне **Параметры** в меню " **Сервис** ").


```vb
Sub PubPath() 
 Dim strPubPath 
 strPubPath = Options.PathForPublications 
 MsgBox strPubPath 
End Sub
```


