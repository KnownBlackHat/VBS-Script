Set sapi=CreateObject("SAPI.SpVoice")
say=InputBox("Please Type Here")
sapi.rate=0
sapi.volume=100
sapi.Speak Say
