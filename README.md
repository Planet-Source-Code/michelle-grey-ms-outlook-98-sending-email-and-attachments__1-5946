<div align="center">

## MS Outlook 98 \- Sending Email and Attachments


</div>

### Description

This code will send email through MS Outlook 98, (or the most current version on your computer),

using MS Excel 7.0 or higher.
 
### More Info
 
You must have MS Outlook and MS Excel 7.0 or higher.

This code will open MS Outlook and open a new email message, insert the recipients email address,

the attachment, the subject of the message, and send the email.

**Note: This is with the aid of using John O'Donald's code and revising it to work for me.



Note: The MailItem.Send line may not work in Excel 97.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Michelle Grey](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/michelle-grey.md)
**Level**          |Intermediate
**User Rating**    |3.8 (15 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/michelle-grey-ms-outlook-98-sending-email-and-attachments__1-5946/archive/master.zip)





### Source Code

```

Sub EmailFromOutlookInExcel() 'macro name
  Set myOlApp = CreateObject("Outlook.Application") 'opens Outlook
  Set MailItem = myOlApp.CreateItem(olMailItem)   ' opens new email
  Set myRecipient = MailItem.Recipients.Add("recipient@company.com")  'inserts recipient's email address
  MailItem.Subject = "Subject of message goes here"   'subject of the email
  Set myAttachments = MailItem.Attachments.Add("C:\foldername\filename")  'Path to Attachments
'Repeat this line if there are more Attachments
  MailItem.Send  'sends the email
End Sub
```

