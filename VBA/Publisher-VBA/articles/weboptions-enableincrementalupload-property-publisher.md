---
title: "Свойство WebOptions.EnableIncrementalUpload (издатель)"
keywords: vbapb10.chm8257541
f1_keywords: vbapb10.chm8257541
ms.prod: publisher
api_name: Publisher.WebOptions.EnableIncrementalUpload
ms.assetid: 42d5e22e-7e39-848e-a7e7-5d2019f7e71c
ms.date: 06/08/2017
ms.openlocfilehash: 13867c84d700cd21de42182bf12d560ee29bae64
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="weboptionsenableincrementalupload-property-publisher"></a>Свойство WebOptions.EnableIncrementalUpload (издатель)

Возвращает или задает **логическое** значение, указывающее, могут быть загружены изменений, внесенных в веб-публикации на веб-сервере, вне зависимости от всей публикации. Если **значение True**, только изменения, внесенные в публикацию будет отправлен на веб-сервер при публикации. Если **значение False**, всей публикации будут отправлены в веб-сервере. Значение по умолчанию — **True**. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **EnableIncrementalUpload**

 переменная _expression_A, представляющий объект **WebOptions** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Свойство **EnableIncrementalUpload** применяется только к веб-публикации, которые уже были опубликованы на веб-сервере. Если веб-публикации еще не был опубликован на веб-сервере, будет опубликована всей публикации на сервере во время первоначальной публикации, независимо от того, является ли свойство **EnableIncrementalUpload** имеет значение **True**. Если веб-публикации еще не были опубликованы в веб-сервера и свойство **EnableIncrementalUpload** затем задано значение **True**, только изменения, внесенные в веб-публикации, а не всю публикацию, после этого момента будет опубликована на сервере.


## <a name="example"></a>Пример

Следующий пример проверяет, является ли веб-публикации задано значение Отправка только изменения, внесенные в публикации. В противном случае свойство **EnableIncrementalUpload** имеет значение **True** для указания, что только изменения публикации будут загружаться на веб-сервере.


```vb
Dim theWO As WebOptions 
 
Set theWO = Application.WebOptions 
 
With theWO 
 If .EnableIncrementalUpload = False Then 
 .EnableIncrementalUpload = True 
 End If 
End With
```


