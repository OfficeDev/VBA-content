---
title: "Объект WebOptions (издатель)"
keywords: vbapb10.chm8323071
f1_keywords: vbapb10.chm8323071
ms.prod: publisher
api_name: Publisher.WebOptions
ms.assetid: 15358c46-f7ca-bc37-d7ef-7d4dbfee09a4
ms.date: 06/08/2017
ms.openlocfilehash: 4af1ab2bdca66114fb93d6dd26c42795b6dc65ed
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="weboptions-object-publisher"></a>Объект WebOptions (издатель)

Представляет свойства веб-публикации, включая параметры для сохранения и кодирования публикации, а также включение безопасные шрифты и схемы шрифтов. Объект **WebOptions** , является участником объекта **[приложения](application-object-publisher.md)** .
 


## <a name="remarks"></a>Заметки

Свойства объекта **WebOptions** используются для указания режима веб-публикации. Это означает, что если какие-либо из этих свойств изменяются, только что созданный веб-публикации наследовать измененных свойств.
 

 
Обратите внимание, что объект **WebOptions** от публикаций и веб-публикации. Тем не менее свойства этого объекта не действовать на печатных публикаций.
 

 

## <a name="example"></a>Пример

Свойство **[WebOptions](application-weboptions-property-publisher.md)** объекта **приложения** используется для возврата объекта **WebOptions** . В следующем примере задается значение объекта Microsoft Publisher **WebOptions** объектную переменную.
 

 

```
Dim theWO As WebOptions 
 
Set theWO = Application.WebOptions
```


## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[AlwaysSaveInDefaultEncoding](weboptions-alwayssaveindefaultencoding-property-publisher.md)|
|[Приложения](weboptions-application-property-publisher.md)|
|[EmailAsImg](weboptions-emailasimg-property-publisher.md)|
|[EnableIncrementalUpload](weboptions-enableincrementalupload-property-publisher.md)|
|[Кодировка](weboptions-encoding-property-publisher.md)|
|[OrganizeInFolder](weboptions-organizeinfolder-property-publisher.md)|
|[Родительский раздел](weboptions-parent-property-publisher.md)|
|[RelyOnVML](weboptions-relyonvml-property-publisher.md)|
|[ShowOnlyWebFonts](weboptions-showonlywebfonts-property-publisher.md)|

