---
title: "Свойство WebOptions.AlwaysSaveInDefaultEncoding (издатель)"
keywords: vbapb10.chm8257539
f1_keywords: vbapb10.chm8257539
ms.prod: publisher
api_name: Publisher.WebOptions.AlwaysSaveInDefaultEncoding
ms.assetid: e37ff08f-5c09-0a71-27e1-e2a332147087
ms.date: 06/08/2017
ms.openlocfilehash: 905763a62b188749810631cb9477be7bbe0e99d8
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="weboptionsalwayssaveindefaultencoding-property-publisher"></a>Свойство WebOptions.AlwaysSaveInDefaultEncoding (издатель)

Возвращает или задает **логическое** значение, указывающее, является ли веб-страницы в веб-публикации всегда сохраняйте с использованием кодировки по умолчанию. Если **значение True**, веб-страниц в пределах публикации всегда будут сохраняться с использованием кодировки по умолчанию на клиентском компьютере. Если **значение False**, веб-страницы не будут сохранены с использованием кодировки по умолчанию. Значение по умолчанию — **False**. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **AlwaysSaveInDefaultEncoding**

 переменная _expression_A, представляет собой объект- **WebOptions** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Если свойство **AlwaysSaveInDefaultEncoding** имеет значение **True** для того или иного объекта **WebOptions** , будут игнорироваться все последующие попытки присвойте свойству **[Кодировка](weboptions-encoding-property-publisher.md)** для этого объекта.


## <a name="example"></a>Пример

В следующем примере проверяется ли веб-публикации в настоящее время задано значение быть сохранен в кодировке по умолчанию. Если так, свойство **AlwaysSaveInDefaultEncoding** имеет значение **False**, а свойство **Encoding** используется для задания кодировки Юникод (UTF-8).


```vb
Dim theWO As WebOptions 
 
Set theWO = Application.WebOptions 
 
With theWO 
 If .AlwaysSaveInDefaultEncoding = True Then 
 .AlwaysSaveInDefaultEncoding = False 
 .Encoding = msoEncodingUTF8 
 End If 
End With
```


