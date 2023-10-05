import pandas as pd
import requests
import time
import pickle
import pygame.mixer
import datetime
import re  # <- Import regular expressions


def timedelta_to_str(td):
    """Convert a timedelta object to a string representation."""
    days, remainder = divmod(
        td.total_seconds(), 86400)  # 86400 seconds in a day
    hours, remainder = divmod(remainder, 3600)  # 3600 seconds in an hour
    minutes, seconds = divmod(remainder, 60)

    parts = []
    if days > 0:
        parts.append(f"{int(days)}d")
    if hours > 0:
        parts.append(f"{int(hours):02}h")
    # Always show minutes and seconds in 2-digit format
    parts.append(f"{int(minutes):02}m {int(seconds):02}s")

    return ' '.join(parts)


def should_skip_translation(text):
    """Check if the text matches certain patterns that should not be translated."""
    if not isinstance(text, str):
        text = str(text)
    pattern = re.compile(r'^\d+\sV$')
    if pattern.match(text):
        return True
    return False


def col_num_to_letter(col_num):
    """Convert a column number into its Excel-style column letter equivalent."""
    letter = ''
    while col_num:
        col_num, remainder = divmod(col_num - 1, 26)
        letter = chr(65 + remainder) + letter
    return letter


# Load your Excel file
df = pd.read_excel('translate.xlsx', engine='openpyxl')

# Create a copy of the dataframe for translations
df_translated = df.copy()

# Change the datatype of all columns in df_translated to object
for col in df_translated.columns:
    df_translated[col] = df_translated[col].astype('object')

# Set up Google Cloud Translation API endpoint and key
endpoint = "https://translation.googleapis.com/language/translate/v2"
api_key = "YOUR_API_KEY"  # Make sure to replace this with your actual API key

# Try to load the cache from a file
try:
    with open('translation_cache.pkl', 'rb') as f:
        translation_cache = pickle.load(f)
except FileNotFoundError:
    translation_cache = {}  # If file not found, initialize an empty cache

# Get total number of rows and columns in the dataframe
total_rows, total_columns = df.shape

# Track start time
start_time = datetime.datetime.now()

# Iterate over each cell in the dataframe, skipping the first row and first column
for row_idx in range(1, total_rows):
    for col_idx in range(1, total_columns):
        cell_content = df.iat[row_idx, col_idx]

        if pd.isna(cell_content) or cell_content == '':
            continue

        if should_skip_translation(str(cell_content)):
            continue

        if cell_content in translation_cache:
            translation = translation_cache[cell_content]
        else:
            data = {
                'q': cell_content,
                'source': 'source_language',
                'target': 'target_language',
                'key': api_key
            }
            response = requests.post(endpoint, data=data)
            translation = response.json(
            )['data']['translations'][0]['translatedText']

            # Save the translation to the cache
            translation_cache[cell_content] = translation

            # Sleep to prevent hitting rate limits
            time.sleep(0.1)

        # Assign the translated content to the appropriate cell in df_translated
        df_translated.iat[row_idx, col_idx] = translation

        # Calculate and print status
        cells_processed = (row_idx * total_columns) + col_idx
        total_cells = total_rows * total_columns
        percentage_done = (cells_processed / total_cells) * 100
        elapsed_time = datetime.datetime.now() - start_time
        estimated_time_left = (elapsed_time / cells_processed) * \
            (total_cells - cells_processed)
        estimated_time_str = timedelta_to_str(estimated_time_left)

        # Display Excel-style cell coordinates
        excel_col = col_num_to_letter(col_idx + 1)
        excel_row = row_idx + 1
        print(
            f"\rTranslating cell {row_idx},{col_idx} ({excel_col}{excel_row}). Progress: {percentage_done:.2f}%. Estimated time left: {estimated_time_str}  ", end='')

# Save the new Excel file
df_translated.to_excel('your_output_file.xlsx', index=False)

# At the end of the script, save the cache to the file
with open('translation_cache.pkl', 'wb') as f:
    pickle.dump(translation_cache, f)

# Play sound at the end
pygame.mixer.init()
sound = pygame.mixer.Sound('src/belldone.mp3')
sound.play()
# Added newline before this print to move to next line after the status updates
print("\nScript is done.")
time.sleep(sound.get_length())
pygame.mixer.quit()
