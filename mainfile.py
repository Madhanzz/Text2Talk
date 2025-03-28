import pyttsx3
from PyPDF2 import PdfReader
from docx import Document
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import speech_recognition as sr


engine = pyttsx3.init()


is_paused = False
stop_speech = False


engine_lock = threading.Lock()


def read_text_aloud(text, speed, volume, voice_id):
    global is_paused, stop_speech

    with engine_lock:  
        engine.setProperty('rate', speed)  
        engine.setProperty('volume', volume) 
        voices = engine.getProperty('voices')
        engine.setProperty('voice', voices[voice_id].id)  

   
    for line in text.split('\n'):
        if stop_speech:
            break  
        while is_paused:
            engine.stop()
            with engine_lock:
                engine.runAndWait()  
        with engine_lock:
            engine.say(line)
            engine.runAndWait()

    
    is_paused = False
    stop_speech = False


def extract_text_from_pdf(file_path):
    reader = PdfReader(file_path)
    text = ''
    for page in reader.pages:
        text += page.extract_text()
    return text

def extract_text_from_docx(file_path):
    doc = Document(file_path)
    text = ''
    for paragraph in doc.paragraphs:
        text += paragraph.text + '\n'
    return text


def start_reading_thread():
    global stop_speech
    stop_speech = False  
    file_path = file_path_var.get()
    if not file_path:
        messagebox.showerror("Error", "No file selected!")
        return

    try:
        if file_path.endswith('.pdf'):
            text = extract_text_from_pdf(file_path)
        elif file_path.endswith('.docx'):
            text = extract_text_from_docx(file_path)
        elif file_path.endswith('.txt'):
            with open(file_path, 'r') as file:
                text = file.read()
        else:
            messagebox.showerror("Error", "Unsupported file format!")
            return

      
        speed = int(speed_var.get())
        volume = float(volume_var.get())
        voice_id = voice_mapping[voice_var.get()]  

        
        threading.Thread(target=read_text_aloud, args=(text, speed, volume, voice_id), daemon=True).start()
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read the document: {e}")


def pause_reading():
    global is_paused
    is_paused = True


def resume_reading():
    global is_paused
    is_paused = False


def stop_reading():
    global stop_speech, is_paused
    stop_speech = True
    is_paused = False
    with engine_lock:
        engine.stop()  


def voice_command_listener():
    recognizer = sr.Recognizer()
    microphone = sr.Microphone()
    

    while True:
        with microphone as source:
            print("Listening for voice commands...")
            try:
                audio = recognizer.listen(source)
                command = recognizer.recognize_google(audio).lower()
                print(f"Command received: {command}")

                if "pause" in command:
                    pause_reading()
                elif "resume" in command:
                    resume_reading()
                elif "stop" in command:
                    stop_reading()
                elif "start" in command:
                    start_reading_thread()
                else:
                    print("Unknown command.")
            except sr.UnknownValueError:
                print("Could not understand the command. Please try again.")
            except sr.RequestError as e:
                print(f"Error with the recognition service: {e}")


root = tk.Tk()
root.title("Document Reader")

file_path_var = tk.StringVar()
speed_var = tk.StringVar(value="200")
volume_var = tk.StringVar(value="1.0")
voice_var = tk.StringVar(value="Male Voice")


voice_mapping = {"Male Voice": 0, "Female Voice": 1}


tk.Label(root, text="Select a Document:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
tk.Entry(root, textvariable=file_path_var, width=50).grid(row=0, column=1, padx=10, pady=5)
tk.Button(root, text="Browse", command=lambda: file_path_var.set(filedialog.askopenfilename(filetypes=[("All Files", "*.pdf *.docx *.txt")]))).grid(row=0, column=2, padx=10, pady=5)


tk.Label(root, text="Reading Speed (Words per Minute):").grid(row=1, column=0, padx=10, pady=5, sticky="w")
tk.Entry(root, textvariable=speed_var).grid(row=1, column=1, padx=10, pady=5)


tk.Label(root, text="Volume (0.0 to 1.0):").grid(row=2, column=0, padx=10, pady=5, sticky="w")
tk.Entry(root, textvariable=volume_var).grid(row=2, column=1, padx=10, pady=5)


tk.Label(root, text="Select Voice:").grid(row=3, column=0, padx=10, pady=5, sticky="w")
voice_options = tk.OptionMenu(root, voice_var, *voice_mapping.keys())
voice_options.grid(row=3, column=1, padx=10, pady=5)


tk.Button(root, text="Start Reading", command=start_reading_thread, bg="green", fg="white").grid(row=4, column=0, padx=5, pady=10)
tk.Button(root, text="Stop", command=stop_reading, bg="red", fg="white").grid(row=4, column=1, padx=5, pady=10)


threading.Thread(target=voice_command_listener, daemon=True).start()

root.mainloop()
