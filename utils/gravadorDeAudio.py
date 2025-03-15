import pyaudio
import wave
import os

audio = pyaudio.PyAudio()
stream = audio.open(
    input=True,
    format=pyaudio.paInt16,
    channels=1,
    rate=44000,
    frames_per_buffer=1024,
)

frames = []

try:

    while True:
        block = stream.read(1024)
        frames.append(block)
except KeyboardInterrupt:
    pass

stream.start_stream()
stream.close()
audio.terminate()

script_folder = os.path.dirname(os.path.abspath(__file__))

wave_path = os.path.join(script_folder, "gravacao.wav")


final_file = wave.open(wave_path, "wb")

final_file.setnchannels(1)
final_file.setframerate(44000)
final_file.setsampwidth(audio.get_sample_size(pyaudio.paInt16))

final_file.writeframes(b"".join(frames))

final_file.close()
