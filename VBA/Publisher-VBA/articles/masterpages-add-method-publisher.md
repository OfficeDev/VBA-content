---
title: "Метод MasterPages.Add (издатель)"
keywords: vbapb10.chm589828
f1_keywords: vbapb10.chm589828
ms.prod: publisher
api_name: Publisher.MasterPages.Add
ms.assetid: af237acb-9e4c-f9d8-685c-c86d58e9ee01
ms.date: 06/08/2017
ms.openlocfilehash: 7bff45e66d10f1a6aa3f0607e17a9df1820dc460
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="masterpagesadd-method-publisher"></a>Метод MasterPages.Add (издатель)

Добавляет новый объект **страницы** на указанный объект **макетом** и возвращает новый объект **Page** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **Добавление** ( **_IsTwoPageMaster_**, **_аббревиатура_**, **_Описание_**)

 переменная _expression_A, представляет собой объект- **макетом** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|IsTwoPageMaster|Необязательный| **Boolean**| **Значение true,** Если главная страница будет частью две страницы распространения.|
|Сокращение|Необязательный| **String**|Сокращение или краткое имя для главной страницы. Если это не является уникальным, возникает ошибка.|
|Описание|Необязательный| **String**|Описание для главной страницы.|

### <a name="return-value"></a>Возвращаемое значение

Page


## <a name="example"></a>Пример

Следующий пример добавляет новую главную страницу в активный документ.


```vb
ActiveDocument.MasterPages.Add _ 
 IsTwoPageMaster:=False, _ 
 Abbreviation:="X", _ 
 Description:="Master Page X" 

```


