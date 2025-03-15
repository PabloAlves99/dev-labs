import pyaudio
import wave
import os

audio = pyaudio.PyAudio()

# Abrindo o stream de áudio para gravação
stream = audio.open(
    input=True,
    format=pyaudio.paInt16,
    channels=1,
    rate=44000,
    frames_per_buffer=1024,
)

frames = []

try:
    # Loop para gravar continuamente até que o usuário pressione Ctrl+C
    while True:
        block = stream.read(1024)
        frames.append(block)
except KeyboardInterrupt:
    print("Gravação interrompida.")

stream.stop_stream()
stream.close()
audio.terminate()

# Definindo o caminho para salvar o arquivo
script_folder = os.path.dirname(os.path.abspath(__file__))
wave_path = os.path.join(script_folder, "gravacao.wav")


final_file = wave.open(wave_path, "wb")

final_file.setnchannels(1)
final_file.setframerate(44000)
final_file.setsampwidth(audio.get_sample_size(pyaudio.paInt16))

final_file.writeframes(b"".join(frames))

final_file.close()
