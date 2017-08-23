---
title: "Объект WebHiddenFields (издатель)"
keywords: vbapb10.chm4063231
f1_keywords: vbapb10.chm4063231
ms.prod: publisher
api_name: Publisher.WebHiddenFields
ms.assetid: 8ced4021-fa99-39dd-e880-b9793426871f
ms.date: 06/08/2017
ms.openlocfilehash: 1f25ce2b3cc79f83b1c5cab7ed5d5c144487d55b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webhiddenfields-object-publisher"></a>Объект WebHiddenFields (издатель)

Представляет скрытых полей Web, позволяющих веб-страницы для передачи невидимые данных на веб-сервер при отправке веб-страницы. Объект **WebHiddenFields** позволяет управлять скрытых полей, подключенного к кнопки Отправить.
 


## <a name="example"></a>Пример

Используйте свойство **скрытые поля** для доступа к скрытых полей Web. В этом примере добавляет новый скрытого поля Web новой кнопки Отправить.
 

 

```
Sub CreateActionWebButton() 
 With ActiveDocument.Pages(1).Shapes 
 With .AddWebControl _ 
 (Type:=pbWebControlCommandButton, Left:=150, _ 
 Top:=150, Width:=75, Height:=36).WebCommandButton 
 .ButtonText = "Submit" 
 .ButtonType = pbCommandButtonSubmit 
 .HiddenFields.Add Name:="User", Value:="PowerUser" 
 End With 
 End With 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Добавление](webhiddenfields-add-method-publisher.md)|
|[Delete](webhiddenfields-delete-method-publisher.md)|
|[Элемент](webhiddenfields-item-method-publisher.md)|
|[Name](webhiddenfields-name-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](webhiddenfields-application-property-publisher.md)|
|[Count](webhiddenfields-count-property-publisher.md)|
|[Родительский раздел](webhiddenfields-parent-property-publisher.md)|

