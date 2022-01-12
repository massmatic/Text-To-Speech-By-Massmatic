Dim msg, sapi
msg=InputBox("Enter your text for conversion                                                                                                                                                                                          Created by Amit Malhotra" ,"Simpe Text to Speech" ,"Type Here")
Set sapi=CreateObject("sapi.spvoice")
sapi.Speak msg
msgbox("Your Text has been Spoken: ") + msg
msgbox("Thank You for Using Text to Speech By Amit Malhotra")
