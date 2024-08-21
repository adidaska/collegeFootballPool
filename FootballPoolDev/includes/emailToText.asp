<%
Dim phoneNumber, carrierDomain, emailAddress, messageBody

' Example values from your database
phoneNumber = "8636603782"
carrierDomain = "@txt.att.net" ' Verizon in this example

' Construct the email address for the SMS
emailAddress = phoneNumber & "@" & carrierDomain

' Message body
messageBody = "Reminder: Please make your game selection before the deadline!"

' Create and configure the CDO.Message object to send the email
Dim objMessage
Set objMessage = Server.CreateObject("CDO.Message")
objMessage.Subject = "Game Selection Reminder"
objMessage.From = "jaflpool@gmail.com"
objMessage.To = emailAddress
objMessage.TextBody = messageBody

' Configure the SMTP server (replace with your SMTP server details)
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 ' cdoSendUsingPort
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.yourisp.com"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 ' Default SMTP port

' Send the message
On Error Resume Next
objMessage.Send

If Err.Number <> 0 Then
    Response.Write "Error sending SMS: " & Err.Description
Else
    Response.Write "SMS sent successfully to " & phoneNumber
End If

' Clear the error and clean up
Err.Clear
Set objMessage = Nothing
%>

<!--Common Carrier Domains-->
<!--Here are some common carrier domains for major U.S. carriers:-->

<!--AT&T: [PhoneNumber]@txt.att.net-->
<!--Verizon: [PhoneNumber]@vtext.com-->
<!--T-Mobile: [PhoneNumber]@tmomail.net-->
<!--Sprint: [PhoneNumber]@messaging.sprintpcs.com-->
<!--US Cellular: [PhoneNumber]@email.uscc.net-->
<!--Boost Mobile: [PhoneNumber]@sms.myboostmobile.com-->
<!--Virgin Mobile: [PhoneNumber]@vmobl.com-->