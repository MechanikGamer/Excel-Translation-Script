# Excel Translation Script

This script facilitates the translation of content in an Excel sheet using the Google Cloud Translation API.

## Setup

### API Key

Replace the placeholder "YOUR_API_KEY" in the script with your actual Google Cloud Translation API key.

### Source and Target Language

Modify the following lines in the script to set your source and target languages:

```python
'source': 'source_language',
'target': 'target_language',
```

Replace 'source_language' and 'target_language' with the desired ISO language codes. For example, for English to Spanish translation, use:

```python
'source': 'en',
'target': 'es',
```

### Input Excel File

The script expects an Excel file named translate.xlsx in the same directory. Ensure your data is in this file before running the script.

## 🚀 Usage

To run the script, navigate to its directory and use:

```shell
python3 translate.py
```

## 💡 How It Works

The script will process cells from the Excel file, skipping the first row and first column. The translation progress will be shown in the format:

```shell
Translating cell 45,6 (G46). Progress: 70.5%. Estimated time left: 2h 15m 35s
```

## 🔍 Logic in the Script

| Component                 | Explanation                                                                             |
| ------------------------- | --------------------------------------------------------------------------------------- |
| `timedelta_to_str`        | Converts time deltas to a user-friendly string format.                                  |
| `should_skip_translation` | Determines if a cell's content should be skipped based on specific patterns.            |
| `col_num_to_letter`       | Converts a numeric column index to its corresponding Excel-style letter (e.g., 1 -> A). |
| `translation_cache`       | Stores previously translated content, reducing the number of API calls.                 |
| `pygame.mixer`            | Plays a sound notification when the script completes its execution.                     |

## 📝 Notes:

- Ensure your API limits on the Google Cloud Translation API are sufficient for the number of cells you wish to translate.
- The script uses caching (`translation_cache.pkl`) to save translations.
- This ensures that if you run the script multiple times, already translated content won't be translated again, saving on API calls.
- The output translated file is saved as `your_output_file.xlsx`.

The script uses caching (translation_cache.pkl) to save translations. This ensures that if you run the script multiple times, already translated content won't be translated again, saving on API calls.
The output translated file is saved as your_output_file.xlsx.

## 📜 Changelog
