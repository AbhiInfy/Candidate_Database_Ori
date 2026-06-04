from pathlib import Path

import win32com.client


ROOT = Path(__file__).resolve().parent
OUT = ROOT / "Hiring Automation Tool"
SCRIPT = OUT / "Hiring Automation Tool - Voiceover Script.txt"
WAV = OUT / "Hiring Automation Tool - Narration.wav"


def main():
    text = SCRIPT.read_text(encoding="utf-8")
    voice = win32com.client.Dispatch("SAPI.SpVoice")
    stream = win32com.client.Dispatch("SAPI.SpFileStream")
    # SSFMCreateForWrite = 3
    stream.Open(str(WAV), 3)
    voice.AudioOutputStream = stream
    voice.Rate = -1
    voice.Volume = 100
    voice.Speak(text)
    stream.Close()
    print(WAV)


if __name__ == "__main__":
    main()
