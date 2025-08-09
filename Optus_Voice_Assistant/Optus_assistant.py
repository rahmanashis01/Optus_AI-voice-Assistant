import os
import re
import struct
import sys
import webbrowser
import platform
import subprocess
import threading
import time
import secrets
import string
import pyperclip
import wikipediaapi
import pyjokes
import qrcode
from PIL import Image
from datetime import datetime

# --- Library Imports with Error Handling ---
try:
    import openai
    import pvporcupine
    import pyttsx3
    import pywhatkit
    import speech_recognition as sr
    from pyaudio import PyAudio
    from screen_brightness_control import get_brightness, set_brightness
    from serpapi import GoogleSearch
    from ctypes import cast, POINTER
    from comtypes import CLSCTX_ALL
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
    import speedtest
    import psutil
    import pyautogui
except ImportError as e:
    print(f"A required library is missing: {e}")
    print("Please run 'pip install -r requirements.txt' in your terminal.")
    sys.exit(1)

# --- CONFIGURATION ---

# -- Porcupine Wake Word Engine --
PORCUPINE_ACCESS_KEY = "EFf7gvqU4fgaPejGy/ISsR5FRPUGReS9YeZQ//TKN7lGvoFhbFEdsg=="

# -- DeepSeek API Key --
DEEPSEEK_API_KEY = "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"

# -- SerpApi Key (for real-time data) --
SERP_API_KEY = "af889687e5c9eae0340d779e13f55f503c26cea9813bd079b56a57f45ee7eca2"

# -- Voice Assistant Settings --
ASSISTANT_NAME = "Optus"
WAKE_WORD = "hey optus"
KEYWORD_FILE_PATH = "hey-optus_en_windows_v3_0_0.ppn"
NOTES_FILE = "notes.txt"

# --- INITIALIZATION ---
print("--- Initializing Optus ---")

# -- Initialize clients, with flags to check their status --
client_deepseek = None
client_wiki = None
engine_tts = None

try:
    client_deepseek = openai.OpenAI(api_key=DEEPSEEK_API_KEY, base_url="https://api.deepseek.com/v1", timeout=20.0)
    print("DeepSeek client initialized.")
except Exception as e:
    print(f"Warning: Could not initialize DeepSeek client: {e}")

try:
    engine_tts = pyttsx3.init('sapi5')
    voices = engine_tts.getProperty('voices')
    if voices:
        engine_tts.setProperty('voice', voices[1].id)
        print("TTS engine initialized.")
    else:
        print("Warning: No TTS voices found.")
        engine_tts = None
except Exception as e:
    print(f"Error initializing TTS engine: {e}")

try:
    client_wiki = wikipediaapi.Wikipedia('OptusAssistant/1.0 (merlin@example.com)')
    print("Wikipedia client initialized.")
except Exception as e:
    print(f"Warning: Could not initialize Wikipedia client: {e}")

recognizer = sr.Recognizer()

# -- Porcupine Wake Word Detection --
try:
    if not os.path.exists(KEYWORD_FILE_PATH):
        print(f"ERROR: Wake word file '{KEYWORD_FILE_PATH}' not found.")
        sys.exit(1)
    porcupine = pvporcupine.create(access_key=PORCUPINE_ACCESS_KEY, keyword_paths=[KEYWORD_FILE_PATH])
    pa = PyAudio()
    audio_stream = pa.open(rate=porcupine.sample_rate, channels=1, format=pa.get_format_from_width(2), input=True,
                           frames_per_buffer=porcupine.frame_length)
    print("Porcupine wake word engine ready.")
except Exception as e:
    print(f"FATAL: Error initializing Porcupine: {e}")
    sys.exit(1)


# --- CORE FUNCTIONS ---

def speak(text):
    """Cleans text and converts it to speech."""
    cleaned_text = str(text).replace('*', '').replace('#', '')
    print(f"{ASSISTANT_NAME}: {cleaned_text}")
    if engine_tts:
        try:
            engine_tts.say(cleaned_text)
            engine_tts.runAndWait()
        except Exception as e:
            print(f"--- TTS Error: {e} ---")


def listen():
    """Listens for and recognizes a command."""
    with sr.Microphone() as source:
        print("Listening...")
        recognizer.pause_threshold = 1
        recognizer.adjust_for_ambient_noise(source)
        audio = recognizer.listen(source)
    try:
        print("Recognizing...")
        command = recognizer.recognize_google(audio, language='en-in').lower().strip()
        print(f"You said: {command}")
        return command
    except sr.UnknownValueError:
        speak("Sorry, I didn't catch that.")
        return ""
    except sr.RequestError:
        speak("Sorry, my speech service is down.")
        return ""
    return ""


def ask_deepseek(prompt):
    """Sends a prompt to the DeepSeek API."""
    if not client_deepseek:
        return "My connection to the DeepSeek AI is currently offline."
    speak("Please be patient, I'm thinking about it.")
    try:
        response = client_deepseek.chat.completions.create(
            model="deepseek-chat",
            messages=[{"role": "user", "content": prompt}],
        )
        return response.choices[0].message.content
    except Exception as e:
        print(f"Error with DeepSeek API: {e}")
        return "Sorry, I couldn't connect to my AI service at the moment."


# --- NEW FEATURE FUNCTIONS ---

def good_morning():
    hour = datetime.now().hour
    if 5 <= hour < 12:
        speak("Good morning! I hope you have a great day.")
    elif 12 <= hour < 18:
        speak("Good afternoon! What can I help you with?")
    else:
        speak("Good evening! How can I assist you?")


def get_trending_searches():
    speak("Let me see what's trending on Google.")
    try:
        search = GoogleSearch({"q": "trending searches", "api_key": SERP_API_KEY})
        results = search.get_dict()
        if "trending_searches" in results:
            trends = [trend['query'] for trend in results['trending_searches'][0]['searches'][:5]]
            speak("Here are some of the top trends:")
            for trend in trends:
                speak(trend)
        else:
            speak("I couldn't fetch the trending searches right now.")
    except Exception as e:
        speak("Sorry, I had trouble getting the trends.")
        print(f"Trending search error: {e}")


def get_internet_speed():
    speak("Testing your internet speed. This might take a moment.")
    try:
        st = speedtest.Speedtest()
        st.download()
        st.upload()
        res = st.results.dict()
        download_speed = res["download"] / 10 ** 6
        upload_speed = res["upload"] / 10 ** 6
        speak(
            f"Your download speed is approximately {download_speed:.2f} megabits per second, and your upload speed is {upload_speed:.2f} megabits per second.")
    except Exception as e:
        speak("I'm sorry, I couldn't measure your internet speed.")
        print(f"Speedtest error: {e}")


def get_pc_specs():
    speak("Here are your computer's specifications.")
    try:
        cpu = f"CPU: {platform.processor()}"
        ram_total = f"Total RAM: {psutil.virtual_memory().total / (1024 ** 3):.2f} GB"
        ram_used = f"Used RAM: {psutil.virtual_memory().percent}%"
        speak(cpu)
        speak(ram_total)
        speak(ram_used)
    except Exception as e:
        speak("I had trouble getting the PC specs.")
        print(f"PC specs error: {e}")


def move_mouse(location):
    width, height = pyautogui.size()
    locations = {
        "top left": (0, 0),
        "top right": (width - 1, 0),
        "bottom left": (0, height - 1),
        "bottom right": (width - 1, height - 1),
        "center": (width / 2, height / 2)
    }
    if location in locations:
        speak(f"Moving mouse to the {location}.")
        pyautogui.moveTo(locations[location][0], locations[location][1], duration=0.5)
    else:
        speak("I don't know that location for the mouse.")


def set_volume_percentage(level):
    try:
        if "mute" in level:
            level_val = 0
        else:
            level_val = int(re.search(r'\d+', level).group())

        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))

        if 0 <= level_val <= 100:
            volume.SetMasterVolumeLevelScalar(level_val / 100, None)
            speak(f"Volume set to {level_val} percent.")
        else:
            speak("Please specify a volume level between 0 and 100.")
    except Exception as e:
        speak("I couldn't change the volume.")
        print(f"Set volume error: {e}")


def start_dictation():
    speak("I'm ready to type. Say 'stop typing' to finish.")
    while True:
        text = listen()
        if "stop typing" in text:
            speak("Dictation finished.")
            break
        elif text:
            pyautogui.write(text + ' ')


def open_folder(folder_name):
    speak(f"Opening your {folder_name} folder.")
    path = ""
    if "download" in folder_name:
        path = os.path.join(os.path.expanduser('~'), 'Downloads')
    elif "document" in folder_name:
        path = os.path.join(os.path.expanduser('~'), 'Documents')

    if path and os.path.exists(path):
        subprocess.Popen(f'explorer "{path}"')
    else:
        speak(f"Sorry, I couldn't find the {folder_name} folder.")


def lock_screen():
    speak("Locking the screen.")
    try:
        import ctypes
        ctypes.windll.user32.LockWorkStation()
    except Exception as e:
        speak("I couldn't lock the screen.")
        print(f"Lock screen error: {e}")


def send_email():
    speak("Opening your default email client for you to send an email.")
    webbrowser.open('mailto:')


def generate_password():
    speak("Generating a secure password. I've copied it to your clipboard.")
    chars = string.ascii_letters + string.digits + string.punctuation
    password = ''.join(secrets.choice(chars) for i in range(16))
    pyperclip.copy(password)
    speak("A new 16-character password is on your clipboard.")


def create_qr_code(data):
    speak("Generating a QR code.")
    try:
        img = qrcode.make(data)
        img_path = "qrcode.png"
        img.save(img_path)

        if platform.system() == 'Windows':
            os.startfile(img_path)
        else:
            subprocess.call(['open', img_path])
        speak("I've created the QR code and opened it for you.")
    except Exception as e:
        speak("I couldn't create the QR code.")
        print(f"QR code error: {e}")


def set_timer(duration_str):
    try:
        minutes = int(re.search(r'\d+', duration_str).group())
        speak(f"Timer set for {minutes} minutes.")

        def timer_end():
            speak(f"Time's up! Your {minutes} minute timer has finished.")

        t = threading.Timer(minutes * 60, timer_end)
        t.daemon = True  # Allows main program to exit even if timer is running
        t.start()
    except Exception as e:
        speak("I couldn't set the timer. Please specify the duration in minutes.")
        print(f"Timer error: {e}")


def add_note():
    speak("What should I write down?")
    note = listen()
    if note:
        with open(NOTES_FILE, 'a') as f:
            f.write(f"- {note} ({datetime.now().strftime('%Y-%m-%d %H:%M')})\n")
        speak("I've added that to your notes.")


def read_notes():
    if os.path.exists(NOTES_FILE):
        speak("Here are your notes.")
        with open(NOTES_FILE, 'r') as f:
            notes = f.read()
        if notes:
            speak(notes)
        else:
            speak("Your notes file is empty.")
    else:
        speak("You don't have any notes yet.")


def calculate(expression):
    try:
        expression = expression.replace("what is", "").replace("calculate", "").strip()
        expression = expression.replace("plus", "+").replace("minus", "-").replace("times", "*").replace("divided by",
                                                                                                         "/")

        allowed_chars = "0123456789+-*/.() "
        if all(char in allowed_chars for char in expression):
            result = eval(expression)
            speak(f"The answer is {result}")
        else:
            speak("I can only handle simple calculations.")
    except Exception as e:
        speak("I couldn't calculate that.")
        print(f"Calculation error: {e}")


def get_wikipedia_summary(command):
    if not client_wiki:
        speak("My connection to Wikipedia is offline.")
        return
    try:
        topic = re.sub(r'wikipedia|tell me about', '', command, flags=re.IGNORECASE).strip()
        speak(f"Getting a summary about {topic} from Wikipedia.")
        page = client_wiki.page(topic)
        if page.exists():
            speak(page.summary[0:500] + "...")
        else:
            speak(f"I couldn't find a Wikipedia page for {topic}.")
    except Exception as e:
        speak("Sorry, I had trouble with Wikipedia.")
        print(f"Wikipedia error: {e}")


def get_news():
    speak("Fetching the latest news headlines.")
    try:
        search = GoogleSearch({"q": "top news headlines", "tbm": "nws", "api_key": SERP_API_KEY})
        results = search.get_dict()
        if "news_results" in results:
            headlines = [res['title'] for res in results['news_results'][:3]]
            for headline in headlines:
                speak(headline)
        else:
            speak("I couldn't fetch the news right now.")
    except Exception as e:
        speak("Sorry, I had trouble getting the news.")
        print(f"News error: {e}")


# --- MAIN COMMAND PROCESSING ---

def process_command(command):
    """Processes the user's command and performs the corresponding action."""
    if not command:
        return

    # --- Daily Routines & Productivity ---
    if "good morning" in command:
        good_morning()
    elif "trending" in command:
        get_trending_searches()
    elif "send an email" in command:
        send_email()
    elif "generate a secure password" in command:
        generate_password()
    elif "make a qr code" in command:
        create_qr_code(command.replace("make a qr code for", "").strip())
    elif "set a timer" in command:
        set_timer(command)
    elif "add to my notes" in command or "take a note" in command:
        add_note()
    elif "read my notes" in command:
        read_notes()

    # --- System & PC Control ---
    elif "internet speed" in command:
        get_internet_speed()
    elif "computer specs" in command or "about this computer" in command:
        get_pc_specs()
    elif "move the mouse" in command:
        move_mouse(command.replace("move the mouse to the", "").strip())
    elif "set volume" in command or "mute volume" in command:
        set_volume_percentage(command)
    elif "start typing" in command:
        start_dictation()
    elif "open my" in command and "folder" in command:
        open_folder(command.replace("open my", "").strip())
    elif "lock the screen" in command or "lock computer" in command:
        lock_screen()
    elif "shut down the computer" in command:
        speak("Are you sure? This will shut down the computer immediately.")
        if "yes" in listen(): os.system("shutdown /s /t 1")
    elif "restart the computer" in command:
        speak("Are you sure? This will restart the computer immediately.")
        if "yes" in listen(): os.system("shutdown /r /t 1")

    # --- Information & Search ---
    elif "calculate" in command or (
            "what is" in command and any(op in command for op in ["times", "plus", "minus", "divided by"])):
        calculate(command)
    elif "what time is it" in command:
        speak(f"The current time is {datetime.now().strftime('%I:%M %p')}.")
    elif "tell me about the movie" in command:
        movie_name = command.replace("tell me about the movie", "").strip()
        speak(ask_deepseek(f"Tell me about the movie {movie_name}. Keep it to a short paragraph."))
    elif "wikipedia" in command:
        get_wikipedia_summary(command)
    elif "define the word" in command:
        speak(ask_deepseek(f"Define the word '{command.replace('define the word', '').strip()}' in a single sentence."))
    elif "search for" in command:
        query = command.replace("search for", "").strip()
        speak(f"Here's what I found for {query}.")
        try:
            pywhatkit.search(query)
        except Exception:
            webbrowser.open(f"https://google.com/search?q={query}")
    elif "news" in command or "headlines" in command:
        get_news()
    elif "give me a quote" in command:
        speak(ask_deepseek("Tell me an inspiring quote."))

    # --- Fun & Media ---
    elif "tell me a joke" in command:
        speak(pyjokes.get_joke())
    elif "play" in command and "on youtube" in command:
        pywhatkit.playonyt(command.replace("on youtube", "").strip())
    elif "open whatsapp" in command:
        try:
            os.startfile("whatsapp:")
            speak("Opening WhatsApp.")
        except Exception:
            speak("I couldn't open the WhatsApp app, opening it in your browser instead.")
            webbrowser.open("https://web.whatsapp.com")
    elif "open browser" in command or "open chrome" in command:
        speak("Opening your browser.")
        webbrowser.open("http://google.com")

    # --- Assistant Control ---
    elif "stop jarvis" in command or "stop optus" in command or "exit program" in command:
        speak("Goodbye!")
        sys.exit()

    # --- Fallback to DeepSeek ---
    else:
        speak(ask_deepseek(command))


# --- MAIN LOOP ---

def main():
    """The main function that runs the voice assistant."""
    speak(f"{ASSISTANT_NAME} is ready. Say '{WAKE_WORD}' to activate.")
    while True:
        try:
            pcm = audio_stream.read(porcupine.frame_length, exception_on_overflow=False)
            pcm = struct.unpack_from("h" * porcupine.frame_length, pcm)
            keyword_index = porcupine.process(pcm)

            if keyword_index >= 0:
                speak("Yes?")
                command = listen()

                if command:
                    if "wait" in command or "stop" in command and len(command) < 6:
                        speak("Okay, I'll wait.")
                        continue
                    process_command(command)

        except KeyboardInterrupt:
            print("\nShutting down...")
            break
        except Exception as e:
            print(f"An error occurred in the main loop: {e}")
            time.sleep(1)


if __name__ == "__main__":
    main()
