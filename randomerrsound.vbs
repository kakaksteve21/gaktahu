Set oVoice = CreateObject("SAPI.SpVoice")
set oSpFileStream = CreateObject("SAPI.SpFileStream")
oSpFileStream.Open "C:\Windows\Media\Windows Critical Stop.wav"
oVoice.SpeakStream oSpFileStream
oSpFileStream.Close
Set oVoice = CreateObject("SAPI.SpVoice")
set oSpFileStream = CreateObject("SAPI.SpFileStream")
oSpFileStream.Open "C:\Windows\Media\Windows Ding.wav"
oVoice.SpeakStream oSpFileStream
oSpFileStream.Close
Set oVoice = CreateObject("SAPI.SpVoice")
set oSpFileStream = CreateObject("SAPI.SpFileStream")
oSpFileStream.Open "C:\Windows\Media\Windows Error.wav"
oVoice.SpeakStream oSpFileStream
oSpFileStream.Close
Set oVoice = CreateObject("SAPI.SpVoice")
set oSpFileStream = CreateObject("SAPI.SpFileStream")
oSpFileStream.Open "C:\Windows\Media\Windows Exclamation.wav"
oVoice.SpeakStream oSpFileStream
oSpFileStream.Close