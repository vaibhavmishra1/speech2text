import speech_recognition as sr 
import win32com.client

import sys,pyperclip
r=sr.Recognizer()
while True:
	k=""
	with sr.Microphone() as source:
		print("say something ")
		audio=r.listen(source)

	try:
		k=r.recognize_google(audio)
		print("you said -",k)
	except:
		pass

	pyperclip.copy(k)
	pyperclip.paste()
	shell=win32com.client.Dispatch("Wscript.Shell")
	shell.SendKeys(k)

