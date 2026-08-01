import json
import os
import tempfile

from sarvamai import SarvamAI


MODEL = "saaras:v4"


def transcribe(audio_path, client, language="en-IN"):

    job = client.speech_to_text_job.create_job(
        model=MODEL,
        mode="transcribe",
        language_code=language
    )

    job.upload_files([audio_path])

    job.start()

    job.wait_until_complete()
    
    print(job)
    print(vars(job))

    with tempfile.TemporaryDirectory() as temp_dir:

        job.download_outputs(temp_dir)

        json_files = [
            f for f in os.listdir(temp_dir)
            if f.endswith(".json")
        ]

        if not json_files:
            raise Exception("No transcript JSON downloaded.")

        json_path = os.path.join(
            temp_dir,
            json_files[0]
        )

        with open(json_path, "r", encoding="utf-8") as f:

            data = json.load(f)

        return data["transcript"]