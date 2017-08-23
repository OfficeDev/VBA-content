---
title: "Объект Printer (издатель)"
keywords: vbapb10.chm9043967
f1_keywords: vbapb10.chm9043967
ms.prod: publisher
api_name: Publisher.Printer
ms.assetid: 46f8c6a2-4cf1-bb6a-1214-a751440870f2
ms.date: 06/08/2017
ms.openlocfilehash: 54350c41858c4eddec2192e46c5efe9b9cc88a2e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="printer-object-publisher"></a>Объект Printer (издатель)

Объект **Printer** представляет принтер, установленный на компьютере.


## <a name="remarks"></a>Заметки

Многие из свойств, таких как **PaperSize**, **PaperSource**и **PaperOrientation**, объект **Printer** соответствуют параметрам в диалоговое окно " **Настройка печати** " (меню " **файл** ") в интерфейсе пользователя Microsoft Publisher.

Коллекция **InstalledPrinters** представлены коллекцию всех принтеров, установленных на вашем компьютере.


## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как можно использовать **Имя_принтера** и **IsActivePrinter** свойства объекта **принтера** , чтобы получить список всех установленных принтеров на компьютере, определить, какой из них в настоящее время активного принтера и получить некоторые параметры активного принтера. Макрос результаты отображаются в окне **Интерпретация** .


```
Public Sub Printer_Example() 
 
 Dim pubInstalledPrinters As Publisher.InstalledPrinters 
 Dim pubApplication As Publisher.Application 
 Dim pubPrinter As Publisher.Printer 
 
 Set pubApplication = ThisDocument.Application 
 Set pubInstalledPrinters = pubApplication.InstalledPrinters 
 
 For Each pubPrinter In pubInstalledPrinters 
 Debug.Print pubPrinter.PrinterName 
 If pubPrinter.IsActivePrinter Then 
 Debug.Print "This is the active printer" 
 Debug.Print "Paper size is ", pubPrinter.PaperSize 
 Debug.Print "Paper orientation is ", pubPrinter.PaperOrientation 
 Debug.Print "Paper source is ", pubPrinter.PaperSource 
 End If 
 Next 
 
End Sub
```


## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](http://msdn.microsoft.com/library/c7eadef4-8206-7e86-b0fe-3c3fe7d07f25%28Office.15%29.aspx)|
|[DriverType](http://msdn.microsoft.com/library/99c3b4e5-a55a-0f8d-3767-d035d9d6e4df%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/2030a3d4-2e42-679c-6084-7a3959271e58%28Office.15%29.aspx)|
|[IsActivePrinter](http://msdn.microsoft.com/library/578fc5d4-2601-66db-cdec-657814756e29%28Office.15%29.aspx)|
|[IsColor](http://msdn.microsoft.com/library/ae466c89-8da0-986b-c3f8-b0aea651dffe%28Office.15%29.aspx)|
|[IsDuplex](http://msdn.microsoft.com/library/d39beb76-8a30-5f2d-3f04-016cfac943fa%28Office.15%29.aspx)|
|[PaperHeight](http://msdn.microsoft.com/library/2c97adb8-0a24-c375-6105-375b203d5640%28Office.15%29.aspx)|
|[PaperOrientation](http://msdn.microsoft.com/library/f57986b6-e6c4-7a47-af93-56036d667240%28Office.15%29.aspx)|
|[PaperSize](http://msdn.microsoft.com/library/fa7962fb-3ca0-470a-2337-3193ed0be2aa%28Office.15%29.aspx)|
|[PaperSource](http://msdn.microsoft.com/library/3c3f9007-c1ea-6957-6fa5-b34873e0a17f%28Office.15%29.aspx)|
|[PaperWidth](http://msdn.microsoft.com/library/e2f0392f-56b2-0ccb-c96c-0bccf2bfe0a0%28Office.15%29.aspx)|
|[Родительский раздел](http://msdn.microsoft.com/library/4f8994d4-423e-8cc6-fb8f-50c47659e892%28Office.15%29.aspx)|
|[PrintableRect](http://msdn.microsoft.com/library/9d5b8264-9213-3d89-0613-421a4872c158%28Office.15%29.aspx)|
|[Имя_принтера](http://msdn.microsoft.com/library/6987b89b-a77e-03c5-bd7e-015510034550%28Office.15%29.aspx)|
|[Режим печати](http://msdn.microsoft.com/library/47ca11d1-d058-0f4e-dd22-ec452dafaf1a%28Office.15%29.aspx)|

