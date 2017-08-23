---
title: "Свойство WebOptions.Encoding (издатель)"
keywords: vbapb10.chm8257540
f1_keywords: vbapb10.chm8257540
ms.prod: publisher
api_name: Publisher.WebOptions.Encoding
ms.assetid: 0aad6082-0ee4-3be0-14a0-73e219f254a0
ms.date: 06/08/2017
ms.openlocfilehash: 65c682263e511dbf6ef5f127628dc679bc2e2bc6
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="weboptionsencoding-property-publisher"></a>Свойство WebOptions.Encoding (издатель)

Возвращает константу **MsoEncoding** , указывающее, кодировка веб-публикации. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Кодировка**

 переменная _expression_A, представляющий объект **WebOptions** .


### <a name="return-value"></a>Возвращаемое значение

MsoEncoding


## <a name="remarks"></a>Заметки

Если свойство **[AlwaysSaveInDefaultEncoding](weboptions-alwayssaveindefaultencoding-property-publisher.md)** имеет значение **True** для того или иного объекта **WebOptions** , будут игнорироваться все последующие попытки присвойте свойству **Кодировка** для этого объекта.

При попытке установить свойство **Encoding** на константу **MsoEncoding** , недоступны в клиенте результаты компьютера к ошибке времени выполнения.

Значение свойства **Кодировка** может иметь одно из ** [MsoEncoding](http://msdn.microsoft.com/library/286bed6e-6028-a252-5e4f-b505234d9d34%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.


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


