from sarvamai import SarvamAI
from saaras_stt import transcribe


client = SarvamAI(
    api_subscription_key=input("API Key : ").strip()
)

audio = input("Audio : ").strip().strip('"')

text = transcribe(audio, client)

print("\nTRANSCRIPT\n")
print(text)