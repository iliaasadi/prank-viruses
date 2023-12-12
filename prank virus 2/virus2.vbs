do

StrText = ("You have been hacked")
set objvoice = CreateObject("SAPI.SpVoice")

Objvoice.Rate=-2
Objvoice.SpeaK StrText

loop