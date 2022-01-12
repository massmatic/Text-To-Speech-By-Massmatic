Dim Message, speak
Message=Inputbox("Enter Text","Speak")
set speak=createobject("sapi.spvoice")
speak.speak Message
