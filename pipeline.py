import os
import shutil
import subprocess
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

from openpyxl import Workbook, load_workbook
from faster_whisper import WhisperModel
from sarvamai import SarvamAI
from sarvamai.play import save



# ============================================================
#                       CONSTANTS
# ============================================================

def get_whisper_model():

    print("\nWhisper Model")
    print("1. Small")
    print("2. Medium")
    

    choice = ask_choice(
        "\nEnter choice : ",
        {"1", "2"}
    )

    models = {
        "1": "small",
        "2": "medium",

    }

    return models[choice]

MASTER_EXCEL = "Master_Transcript.xlsx"

MASTER_HEADERS = [
    "File Name",
    "Source Transcript"
]

LANGUAGE_HEADERS = [
    "File Name",
    "Source Transcript",
    "Translation"
]

IGNORE_EXCELS = {
    MASTER_EXCEL,
    "Hindi.xlsx",
    "Telugu.xlsx",
    "Marathi.xlsx",
    "Punjabi.xlsx",
    "Tamil.xlsx",
    "Bengali.xlsx",
    "Kannada.xlsx",
    "Gujarati.xlsx"
}

AUDIO_EXTENSIONS = (
    ".wav",
    ".mp3",
    ".flac",
    ".aac",
    ".m4a",
    ".ogg"
)

LANGUAGES = {
    "1": {
        "name": "Hindi",
        "code": "hi-IN",
        "prefix": "hi"
    },
    "2": {
        "name": "Telugu",
        "code": "te-IN",
        "prefix": "te"
    },
    "3": {
        "name": "Marathi",
        "code": "mr-IN",
        "prefix": "mr"
    },
    "4": {
        "name": "Punjabi",
        "code": "pa-IN",
        "prefix": "pa"
    },
    "5": {
        "name": "Tamil",
        "code": "ta-IN",
        "prefix": "ta"
    },
    "6": {
        "name": "Bengali",
        "code": "bn-IN",
        "prefix": "bn"
    },
    "7": {
        "name": "Kannada",
        "code": "kn-IN",
        "prefix": "kn"
    },
    "8": {
        "name": "Gujarati",
        "code": "gu-IN",
        "prefix": "gu"
    }
 }

TRANSLATION_MODES = {
    "1": {
        "name": "Formal",
        "value": "formal"
    },
    "2": {
        "name": "Classic Colloquial",
        "value": "classic-colloquial"
    },
    "3": {
        "name": "Modern Colloquial",
        "value": "modern-colloquial"
    },
    "4": {
        "name": "Code Mixed",
        "value": "code-mixed"
    }
}

# ============================================================
#                     INPUT HELPERS
# ============================================================

def ask_non_empty(message):

    while True:

        value = input(message).strip()

        if value:
            return value

        print("Cannot be empty.")


def ask_choice(message, choices):

    while True:

        value = input(message).strip()

        if value in choices:
            return value

        print("Invalid choice.")


# ============================================================
#                     USER INPUT
# ============================================================

def get_api_key():

    return ask_non_empty("\nSarvam API Key : ")


def get_input_type():

    print("\nInput Type")
    print("1. Audio")
    print("2. Transcript")

    choice = ask_choice(
        "\nEnter choice : ",
        {"1", "2"}
    )

    return "audio" if choice == "1" else "transcript"


def get_project_folder():

    while True:

        folder = input(
            "\nProject Folder : "
        ).strip().strip('"')

        if os.path.isdir(folder):
            return os.path.abspath(folder)

        print("Folder not found.")


def get_processing_mode():

    print("\nProcessing Mode")
    print("1. Sequential")
    print("2. Parallel")

    choice = ask_choice(
        "\nEnter choice : ",
        {"1", "2"}
    )

    return "sequential" if choice == "1" else "parallel"

def get_translation_mode():

    print("\nTranslation Style\n")

    for key, value in TRANSLATION_MODES.items():
        print(f"{key}. {value['name']}")

    choice = ask_choice(
        "\nEnter choice : ",
        TRANSLATION_MODES.keys()
    )

    return TRANSLATION_MODES[choice]["value"]


def get_target_languages():

    print("\nTarget Languages\n")

    for key, value in LANGUAGES.items():

        print(f"{key}. {value['name']}")

    while True:

        raw = input(
            "\nEnter comma separated values : "
        ).replace(" ", "")

        ids = raw.split(",")

        if len(ids) != len(set(ids)):
            print("Duplicate languages selected.")
            continue

        if all(i in LANGUAGES for i in ids):
            return ids

        print("Invalid selection.")

def get_audio_format():

    print("\nOutput Audio Format")

    print("1. OGG (Recommended for Unity)")
    print("2. MP3")
    print("3. WAV")

    choice = ask_choice(
        "\nEnter choice : ",
        {"1", "2", "3"}
    )

    formats = {
        "1": "ogg",
        "2": "mp3",
        "3": "wav"
    }

    return formats[choice]

def configure_languages(language_ids):

    jobs = []

    for language_id in language_ids:

        lang = LANGUAGES[language_id]

        print(f"\n========== {lang['name']} ==========")

        mode = ask_choice(
            "Mode (T / A / TA) : ",
            {"T", "A", "TA"}
        ).upper()

        speaker = None

        if mode in ("A", "TA"):

            speaker = ask_non_empty(
                "Speaker : "
            )

        jobs.append({

            "name": lang["name"],
            "code": lang["code"],
            "prefix": lang["prefix"],
            "mode": mode,
            "speaker": speaker

        })

    return jobs


# ============================================================
#                  PROJECT HELPERS
# ============================================================

def master_excel(project):

    return os.path.join(
        project,
        MASTER_EXCEL
    )


def language_folder(project, job):

    return os.path.join(
        project,
        job["name"]
    )


def language_excel(project, job):

    return os.path.join(
        language_folder(project, job),
        f"{job['name']}.xlsx"
    )


def audio_folder(project, job):

    return os.path.join(
        language_folder(project, job),
        "Audio"
    )


def project_exists(project):

    return os.path.exists(
        master_excel(project)
    )


def get_project_mode(project):

    if not project_exists(project):
        return "fresh"

    print("\nExisting Project Found")
    print("1. Resume")
    print("2. Start Fresh")

    choice = ask_choice(
        "\nEnter choice : ",
        {"1", "2"}
    )

    if choice == "1":
        return "resume"

    if os.path.exists(master_excel(project)):
        os.remove(master_excel(project))

    for lang in LANGUAGES.values():

        folder = os.path.join(
            project,
            lang["name"]
        )

        if os.path.exists(folder):
            shutil.rmtree(folder)

    return "fresh"


def create_output_structure(project, jobs):

    for job in jobs:

        folder = language_folder(project, job)

        os.makedirs(
            folder,
            exist_ok=True
        )

        if job["mode"] in ("A", "TA"):

            os.makedirs(
                audio_folder(project, job),
                exist_ok=True
            )

        job["folder"] = folder
        job["excel"] = language_excel(project, job)
        job["audio_folder"] = audio_folder(project, job)

# ============================================================
#                     VALIDATION
# ============================================================

def validate_audio_project(project):

    files = [

        f for f in os.listdir(project)

        if f.lower().endswith(AUDIO_EXTENSIONS)

    ]

    if not files:

        raise Exception(
            "No audio files found in project folder."
        )


def find_transcript_excel(project):

    candidates = []

    for file in os.listdir(project):

        if not file.lower().endswith(".xlsx"):
            continue

        if file in IGNORE_EXCELS:
            continue

        candidates.append(
            os.path.join(project, file)
        )

    if len(candidates) == 0:

        raise Exception(
            "Transcript Excel not found."
        )

    if len(candidates) > 1:

        raise Exception(
            "Multiple transcript Excel files found."
        )

    return candidates[0]


def validate_headers(ws, required):

    headers = {}

    for col in range(1, ws.max_column + 1):

        value = ws.cell(
            row=1,
            column=col
        ).value

        if value:
            headers[str(value).strip()] = col

    for item in required:

        if item not in headers:

            raise Exception(
                f"Missing header : {item}"
            )

    return headers


# ============================================================
#                 WHISPER / SARVAM
# ============================================================

def load_whisper(model_name):

    print(f"\nLoading Whisper ({model_name})...")

    model = WhisperModel(
        model_name,
        device="cpu",
        compute_type="int8"
    )

    print("Whisper Ready.")

    return model


def create_sarvam_client(api_key):

    return SarvamAI(
        api_subscription_key=api_key
    )

    
def convert_audio(input_file, output_format):

    if output_format == "wav":
        return input_file

    output_file = os.path.splitext(input_file)[0] + f".{output_format}"

    subprocess.run(
        [
            "ffmpeg",
            "-y",
            "-i",
            input_file,
            output_file
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        check=True
    )

    os.remove(input_file)

    return output_file


# ============================================================
#                  MASTER EXCEL
# ============================================================

def create_master_excel(project):

    path = master_excel(project)

    if os.path.exists(path):
        return

    wb = Workbook()

    ws = wb.active

    ws.append(MASTER_HEADERS)

    wb.save(path)

    wb.close()


def load_master(project):

    wb = load_workbook(
        master_excel(project)
    )

    return wb, wb.active


# ============================================================
#                  AUDIO DISCOVERY
# ============================================================

def get_audio_files(project):

    return sorted([

        os.path.join(project, file)

        for file in os.listdir(project)

        if file.lower().endswith(
            AUDIO_EXTENSIONS
        )

    ])


# ============================================================
#             AUDIO -> MASTER TRANSCRIPT
# ============================================================

def transcribe_audio(project, whisper):

    wb, ws = load_master(project)

    headers = validate_headers(
        ws,
        MASTER_HEADERS
    )

    file_col = headers["File Name"]
    text_col = headers["Source Transcript"]

    completed = {

        ws.cell(
            row=row,
            column=file_col
        ).value

        for row in range(
            2,
            ws.max_row + 1
        )

    }

    audio_files = get_audio_files(project)

    total = len(audio_files)

    skipped = 0

    for index, audio in enumerate(audio_files, 1):

        filename = os.path.basename(audio)

        if filename in completed:

            skipped += 1
            continue

        print(
            f"[{index}/{total}] {filename}"
        )

        segments, info = whisper.transcribe(
            audio,
            beam_size=5,
            vad_filter=True
        )

        transcript = " ".join(

            s.text.strip()

            for s in segments

        )

        ws.append([

            filename,
            transcript

        ])

        wb.save(master_excel(project))

    wb.close()

    print(
        f"Skipped : {skipped}"
    )


# ============================================================
#          TRANSCRIPT -> MASTER TRANSCRIPT
# ============================================================

def import_transcript_excel(project):

    source_excel = find_transcript_excel(
        project
    )

    source = load_workbook(
        source_excel
    )

    source_ws = source.active

    source_headers = validate_headers(

        source_ws,

        MASTER_HEADERS

    )

    master_wb, master_ws = load_master(
        project
    )

    master_headers = validate_headers(

        master_ws,

        MASTER_HEADERS

    )

    existing = {

        master_ws.cell(

            row=row,

            column=master_headers[
                "File Name"
            ]

        ).value

        for row in range(
            2,
            master_ws.max_row + 1
        )

    }

    for row in range(
        2,
        source_ws.max_row + 1
    ):

        filename = source_ws.cell(

            row=row,

            column=source_headers[
                "File Name"
            ]

        ).value

        transcript = source_ws.cell(

            row=row,

            column=source_headers[
                "Source Transcript"
            ]

        ).value

        if filename in existing:
            continue

        master_ws.append([

            filename,

            transcript

        ])

    master_wb.save(
        master_excel(project)
    )

    source.close()
    master_wb.close()

# ============================================================
#              LANGUAGE EXCEL CREATION
# ============================================================

def create_language_excel(job):

    if os.path.exists(job["excel"]):
        return

    wb = Workbook()

    ws = wb.active

    ws.append([
        "File Name",
        "Source Transcript",
        "Translation"
    ])

    wb.save(job["excel"])
    wb.close()


def populate_language_excel(project, job):

    create_language_excel(job)

    master_wb, master_ws = load_master(project)

    lang_wb = load_workbook(job["excel"])
    lang_ws = lang_wb.active

    master_headers = validate_headers(
        master_ws,
        MASTER_HEADERS
    )

    lang_headers = validate_headers(
        lang_ws,
        LANGUAGE_HEADERS
    )

    existing = {

        lang_ws.cell(
            row=row,
            column=lang_headers["File Name"]
        ).value

        for row in range(
            2,
            lang_ws.max_row + 1
        )

    }

    changed = False

    for row in range(
        2,
        master_ws.max_row + 1
    ):

        filename = master_ws.cell(
            row=row,
            column=master_headers["File Name"]
        ).value

        transcript = master_ws.cell(
            row=row,
            column=master_headers["Source Transcript"]
        ).value

        if filename in existing:
            continue

        lang_ws.append([
            filename,
            transcript,
            ""
        ])

        changed = True

    if changed:
        lang_wb.save(job["excel"])

    master_wb.close()
    lang_wb.close()


# ============================================================
#                    TRANSLATION
# ============================================================

def translate_language(client, job):

    if job["mode"] == "A":

        if not os.path.exists(job["excel"]):
            return

    wb = load_workbook(job["excel"])
    ws = wb.active

    headers = validate_headers(
        ws,
        LANGUAGE_HEADERS
    )

    file_col = headers["File Name"]
    text_col = headers["Source Transcript"]
    trans_col = headers["Translation"]

    total = ws.max_row - 1
    skipped = 0
    completed = 0

    for row in range(2, ws.max_row + 1):

        filename = ws.cell(
            row=row,
            column=file_col
        ).value

        english = ws.cell(
            row=row,
            column=text_col
        ).value

        translated = ws.cell(
            row=row,
            column=trans_col
        ).value

        if translated:

            skipped += 1
            continue
        if not english or not str(english).strip():
            skipped += 1
            continue

        print(
            f"[{job['name']}] "
            f"{completed+1}/{total} : "
            f"{filename}"
        )

        try:

            response = client.text.translate(
                    input=english,
                    source_language_code="en-IN",
                    target_language_code=job["code"],
                    model="mayura:v1",
                    mode=job["translation_mode"]
                )

            ws.cell(
                row=row,
                column=trans_col
            ).value = response.translated_text

            wb.save(job["excel"])

            completed += 1

        except Exception as e:

            print(e)

    wb.close()

    print(
        f"{job['name']} "
        f"Completed:{completed} "
        f"Skipped:{skipped}"
    )


# ============================================================
#                       TTS
# ============================================================

def generate_audio(client, job, audio_format):

    if job["mode"] == "T":
        return

    wb = load_workbook(job["excel"])
    ws = wb.active

    headers = validate_headers(
        ws,
        LANGUAGE_HEADERS
    )

    file_col = headers["File Name"]
    trans_col = headers["Translation"]

    total = ws.max_row - 1
    skipped = 0
    completed = 0

    for row in range(2, ws.max_row + 1):

        filename = ws.cell(
            row=row,
            column=file_col
        ).value

        translated = ws.cell(
            row=row,
            column=trans_col
        ).value

        # -------- FIXED A MODE --------

        if not translated:

            if job["mode"] == "A":

                print(
                    f"[{job['name']}] "
                    f"Missing translation : "
                    f"{filename}"
                )

            continue

        base = os.path.splitext(filename)[0]

        temp_output = os.path.join(
            job["audio_folder"],
            f"{job['prefix']}_{base}.wav"
        )

        final_output = os.path.join(
            job["audio_folder"],
            f"{job['prefix']}_{base}.{audio_format}"
        )

        if os.path.exists(final_output):
            skipped += 1
            continue

        print(
            f"[{job['name']}] "
            f"{completed+1}/{total} : "
            f"{filename}"
        )

        try:

            audio = client.text_to_speech.convert(
                text=translated,
                target_language_code=job["code"],
                model="bulbul:v3",
                speaker=job["speaker"]
            )

            save(audio, temp_output)

            final_output = convert_audio(
                temp_output,
                audio_format
            )

            completed += 1

        except Exception as e:
            print(e)
            
def process_language(api_key, project, job, audio_format):

    try:

        client = create_sarvam_client(api_key)

        populate_language_excel(
            project,
            job
        )

        translate_language(
            client,
            job
        )

        generate_audio(
            client,
            job,
            audio_format
        )

        print(
            f"{job['name']} Finished."
        )

    except Exception as e:

        print(
            f"{job['name']} Failed : {e}"
        )


# ============================================================
#                PIPELINE RUNNER
# ============================================================

def run_pipeline(

    api_key,

    project,

    jobs,

    processing_mode,
    
    audio_format
  ):

    if processing_mode == "sequential":

        for job in jobs:

            process_language(

                api_key,

                project,

                job,
                
                audio_format

            )

        return

    with ThreadPoolExecutor(

        max_workers=min(

            len(jobs),

            os.cpu_count() or 4

        )

    ) as executor:

        futures = [

            executor.submit(

                process_language,

                api_key,

                project,

                job,
                
                audio_format

            )

            for job in jobs

        ]

        for future in as_completed(futures):

            try:

                future.result()

            except Exception as e:

                print(e)


# ============================================================
#                        MAIN
# ============================================================

def main():

    input_type = get_input_type()

    project = get_project_folder()

    project_mode = get_project_mode(
        project
    )

    create_master_excel(
        project
    )

    if input_type == "audio":

        validate_audio_project(
            project
        )

    else:

        if not os.path.exists(master_excel(project)):
            find_transcript_excel(
                project
            )

    whisper = None

    if input_type == "audio":

        whisper_model = get_whisper_model()
        whisper = load_whisper(whisper_model)

    if project_mode == "fresh":

        if input_type == "audio":

            transcribe_audio(
                project,
                whisper
            )

    else:

        if os.path.exists(master_excel(project)):
            print("Using existing Master_Transcript.xlsx")
        else:
            import_transcript_excel(
                project
            )

    api_key = input(
        "\nSarvam API Key (Press Enter to skip): "
    ).strip()

    if not api_key:

        print("\nSkipping Translation / TTS.")
        print("\n=================================")
        print(" Pipeline Completed Successfully ")
        print("=================================")

        return

    processing_mode = get_processing_mode()

    translation_mode = get_translation_mode()

    language_ids = get_target_languages()

    jobs = configure_languages(language_ids)

    for job in jobs:
        job["translation_mode"] = translation_mode

    audio_format = get_audio_format()

    create_output_structure(
        project,
        jobs
    )

    run_pipeline(
        api_key,
        project,
        jobs,
        processing_mode,
        audio_format
    )

    print("\n=================================")
    print(" Pipeline Completed Successfully ")
    print("=================================")



if __name__ == "__main__":

    main()

# =================== END OF FILE ===================