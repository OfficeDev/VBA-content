---
title: "Объект WebTextBox (издатель)"
keywords: vbapb10.chm4259839
f1_keywords: vbapb10.chm4259839
ms.prod: publisher
api_name: Publisher.WebTextBox
ms.assetid: 74fde391-734c-6672-dadb-59bc58232c0f
ms.date: 06/08/2017
ms.openlocfilehash: e52b35239218ebff72fe34e01951c1569b4da40a
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webtextbox-object-publisher"></a>Объект WebTextBox (издатель)

Представляет элемент управления текстового поля Web. Объект **WebTextBox** является членом объекта **Shape** .
 


## <a name="example"></a>Пример

Используйте метод **[AddWebControl](shapes-addwebcontrol-method-publisher.md)** для создания новой кнопки параметр Web. Используйте свойство **[WebTextBox](shape-webtextbox-property-publisher.md)** для доступа к поле элемента управления Web текст фигуры. В этом примере создается новое текстовое поле Web, задает текст по умолчанию, указывает, что запись является обязательным и ограничения для записи до 50 символов.
 

 

```
Sub CreateWebTextBox() 
 With ActiveDocument.Pages(1).Shapes 
 With .AddWebControl(Type:=pbWebControlSingleLineTextBox, _ 
 Left:=100, Top:=100, Width:=150, Height:=15).WebTextBox 
 .DefaultText = "Please Enter Your Full Name" 
 .RequiredControl = msoTrue 
 .Limit = 50 
 End With 
 End With 
End Sub
```


## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](webtextbox-application-property-publisher.md)|
|[DefaultText](webtextbox-defaulttext-property-publisher.md)|
|[EchoAsterisks](webtextbox-echoasterisks-property-publisher.md)|
|[Ограничение](webtextbox-limit-property-publisher.md)|
|[Родительский раздел](webtextbox-parent-property-publisher.md)|
|[RequiredControl](webtextbox-requiredcontrol-property-publisher.md)|
|[ReturnDataLabel](webtextbox-returndatalabel-property-publisher.md)|

