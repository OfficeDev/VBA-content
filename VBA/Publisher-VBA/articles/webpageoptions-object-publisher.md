---
title: "Объект WebPageOptions (издатель)"
keywords: vbapb10.chm548863
f1_keywords: vbapb10.chm548863
ms.prod: publisher
api_name: Publisher.WebPageOptions
ms.assetid: 694b56ce-1c2d-8202-25b7-19e55aadb0fd
ms.date: 06/08/2017
ms.openlocfilehash: cd6ed0327a90cdd4f6f3dd66dc1be9611715548d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webpageoptions-object-publisher"></a>Объект WebPageOptions (издатель)

Представляет свойства одного веб-страницы в веб-публикации, включая варианты добавления заголовок и описание страницы, звуковое сопровождение, кроме работы с другими параметрами. Объект **WebPageOptions** , является участником объекта **[Page](page-object-publisher.md)** .
 


## <a name="remarks"></a>Заметки

Обратите внимание, что объект **WebPageOptions** доступна только при активной публикации веб-публикации. При попытке доступа к этому объекту из публикации в возвращается ошибка во время выполнения.
 

 

## <a name="example"></a>Пример

Используйте свойство **[WebPageOptions](page-webpageoptions-property-publisher.md)** на объект **Page** для получения объекта **WebPageOptions** . Используйте свойство **[Description](webpageoptions-description-property-publisher.md)** описание для указанного веб-страницы. В следующем примере задается описание для второй страницы active веб-публикации.
 

 

```
Dim theWPO As WebPageOptions 
 
Set theWPO = ActiveDocument.Pages(2).WebPageOptions 
 
With theWPO 
 .Description = "Company Profile" 
End With
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[SetBackgroundSoundRepeat](webpageoptions-setbackgroundsoundrepeat-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](webpageoptions-application-property-publisher.md)|
|['' Фоновый звук ''](webpageoptions-backgroundsound-property-publisher.md)|
|[BackgroundSoundLoopCount](webpageoptions-backgroundsoundloopcount-property-publisher.md)|
|[BackgroundSoundLoopForever](webpageoptions-backgroundsoundloopforever-property-publisher.md)|
|[Описание](webpageoptions-description-property-publisher.md)|
|[IncludePageOnNewWebNavigationBars](webpageoptions-includepageonnewwebnavigationbars-property-publisher.md)|
|[Ключевые слова](webpageoptions-keywords-property-publisher.md)|
|[Родительский раздел](webpageoptions-parent-property-publisher.md)|
|[PublishFileName](webpageoptions-publishfilename-property-publisher.md)|

