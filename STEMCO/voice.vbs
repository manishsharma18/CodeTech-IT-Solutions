set cv= CreateObject("SAPI.spvoice")
set cv.voice = cv.GetVoices.Item(1)  'female voice 
set cv.voice = cv.GetVoices.Item(0)  'male voice and default  voice is also male voice
cv.speak "HELLO EVERYONE"
WScript.sleep 500
cv.speak "nice to meet you"
WScript.sleep 500
cv.speak "get lost"
