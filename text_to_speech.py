from win32com.client import Dispatch

""" Install necessary package: 
    pip install pypiwin32
 """

speaker = Dispatch("SAPI.SpVoice")

print("Write something. 'q' for Quit: ")
user_input = input()

while user_input != 'q':
    print("Write something. 'q' for Quit: ")
    speaker.Speak(user_input)
    user_input = input()

del speaker
