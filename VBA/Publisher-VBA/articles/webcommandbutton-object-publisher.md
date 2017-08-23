---
title: "Объект WebCommandButton (издатель)"
keywords: vbapb10.chm3997695
f1_keywords: vbapb10.chm3997695
ms.prod: publisher
api_name: Publisher.WebCommandButton
ms.assetid: 86605945-eca1-ab80-1a1a-f8a5977d9282
ms.date: 06/08/2017
ms.openlocfilehash: a8afcf97352a934ab59da02e1b33112cdc92913e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webcommandbutton-object-publisher"></a>Объект WebCommandButton (издатель)

Представляет элемент управления кнопки команды Web. Объект **WebCommandButton** является членом объекта **Shape** .
 


## <a name="example"></a>Пример

Используйте метод **[AddWebControl](shapes-addwebcontrol-method-publisher.md)** для создания новой кнопки команды Web. Используйте свойство **[WebCommandButton](shape-webcommandbutton-property-publisher.md)** для доступа к Web командной кнопки управления фигуры. В этом примере создается кнопки Отправить форму Web и задает путь и имя скрипта для запуска при нажатии кнопки.
 

 

```
Sub CreateActionWebButton() 
 With ActiveDocument.Pages(1).Shapes.AddWebControl _ 
 (Type:=pbWebControlCommandButton, Left:=150, _ 
 Top:=150, Width:=75, Height:=36).WebCommandButton 
 .ButtonText = "Submit" 
 .ButtonType = pbCommandButtonSubmit 
 .ActionURL = "http://www.tailspintoys.com/" _ 
 &amp; "scripts/ispscript.cgi" 
 End With 
End Sub
```


## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[ActionURL](webcommandbutton-actionurl-property-publisher.md)|
|[Приложения](webcommandbutton-application-property-publisher.md)|
|[ButtonText](webcommandbutton-buttontext-property-publisher.md)|
|[ButtonType](webcommandbutton-buttontype-property-publisher.md)|
|[DataFileFormat](webcommandbutton-datafileformat-property-publisher.md)|
|[Имя_файла_данных](webcommandbutton-datafilename-property-publisher.md)|
|[DataRetrievalMethod](webcommandbutton-dataretrievalmethod-property-publisher.md)|
|[EmailAddress](webcommandbutton-emailaddress-property-publisher.md)|
|[EmailSubject](webcommandbutton-emailsubject-property-publisher.md)|
|[Скрытые поля](webcommandbutton-hiddenfields-property-publisher.md)|
|[Родительский раздел](webcommandbutton-parent-property-publisher.md)|
|[PostFormData](webcommandbutton-postformdata-property-publisher.md)|

