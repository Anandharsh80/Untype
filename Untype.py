#Beginning of Script



# UNTYPE
# Code by - HARSH ANAND

# Pre requisites to run the following script:
# speech_recognition, pyaudio, python-docx, termcolor packages must be installed to run this script 


# Importing Speech recognition API and colored sub-package from termcolor
import speech_recognition as voice
from termcolor import colored
from docx import Document




# The Speech recognizer function that takes real time speech as Input
def SpeechRecognizerRealTime():

    # Declaring the voice recognizer variable that will detact and convert the Audio
    v = voice.Recognizer()

    # Using the Device Microphone to take Audio input
    with voice.Microphone() as source:

        print(colored("Talk","blue"))

        # Taking input and storing it
        audio = v.listen(source)
        print(colored("Time over, Thanks","magenta"))

    # Checking for Unknown Value error or any other Exception
    try:
        
        # Recognising the audio on the basis of specified language (here English(INDIA)), and converting it into a string
        Input_Text = v.recognize_google(audio, language = "en-IN")
        Input_Text = Input_Text.capitalize() + ". "
        print(colored("Last Speech: ","magenta") + Input_Text)
        return Input_Text

    # Catching any exception That may have occoured 
    except Exception:
        print("Not able to detect any kind of audio")
        return ""





# The Speech recognizer function that takes real time speech as Input
def SpeechRecognizerSourceReader(source):

    # Declaring the voice recognizer variable that will recognize and convert the Audio
    v = voice.Recognizer()

    with voice.AudioFile(source) as aud:
        # Taking input from the Souce location and storing it
        audio = v.listen(aud)

    # Checking for Unknown Value error or any other Exception
    try:
        
        # Recognising the audio on the basis of specified language (here English(INDIA)), and converting it into a string
        Input_Text = v.recognize_google(audio, language = "en-IN")
        Input_Text = Input_Text.capitalize() + ". "
        print(colored("Last Speech: ","magenta") + Input_Text)
        return Input_Text

    # Catching any exception That may have occoured 
    except Exception:
        print(colored("Not able to detect any kind of audio","red"))
        return ""





# The Document File Generator Function
def WriteDocFile(destination, data):

    # Declaring the Document Variable
    doc = Document()

    doc.add_paragraph(data)
    doc.save(destination)






# Driver Code

if __name__ == "__main__":


    # Speech to Text Recognition part

    print(colored("Welcome to Speech recognizer", "cyan"))
    TotalText = ""

    while (True):

        print(colored("Enter 1 to Speak","green"))
        print(colored("Enter 2 to use a Pre-Recorded File","green"))
        val = int(input())
        if (val==1):
            func_return = SpeechRecognizerRealTime()
        elif (val==2):
            print(colored("Enter recorder file","magenta"))
            source = input()
            func_return = SpeechRecognizerSourceReader(source)
        else: 
            print(colored("Invalid Choice","red"))
            continue

        if (func_return == ""):
            continue
        else:
            print(colored("Are you satisfied or do you wish to try again ??\nEnter '0' to try again and '1' to continue.","green"))
            
            while (True):
                var = int(input())
                if (var==0 or var==1):
                    break
                print(colored("Invalid Choice","red"))
            
            if (var == 0):
                continue
            else:
                TotalText = TotalText + func_return
                TotalText = TotalText + "\n"

        print(colored("Total Text so far :\n", "cyan") + TotalText)
        print(colored("Are you satisfied or do you wish to add more text??\nEnter '0' to add more text and '1' to continue.","green"))
        while (True):
            var = int(input())
            if (var==0 or var==1):
                break
            print(colored("Invalid Choice","red"))
            
        if (var == 0):
            continue
        else:
            break

    FinalText = TotalText
    print(colored(colored("Final Text is as follows: \n","cyan") + colored(FinalText,"yellow")))


    # Text to Document creation Part

    print(colored("Do you wish to convert this text into a Document File ??","magenta"))
    print(colored("Press '1' to continue or Press '0' to skip","cyan"))
    while (True):
        var = int(input())
        if (var==0 or var==1):
            break
        print(colored("Invalid Choice","red"))

    if (var == 1):
        print(colored("Enter the Name of the file to be stored","green"))
        fName = input()
        fName = fName + '.docx'
        WriteDocFile(fName,FinalText)



# End of Script