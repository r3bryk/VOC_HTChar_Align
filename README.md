# VOC_HTChar_Align.py

Version 1.5 (2026-11-06)

## Description

`VOC_HTChar_Align.py` is a script for inter-sample GC-MS data alignment and compound classification.

## What does the script do

The script takes an `Excel` file as input (one sheet per sample) and does the following:

1. Reading and cleaning of the input tables (column names, whitespace, CAS# format, numeric values).
2. Inter-sample alignment based on exact matching of the compound names (case-insensitive) and CAS# and tolerance-based (default: 10 s ≈ 0.1667 min) matching of retention times.
3. Retrieval of InChIKey values from the PubChem database (Kim et al. 2023. Nucleic Acids Res.) based on the compound names, with the CAS# as a fallback.
4. Classification of the compounds using ClassyFire (Djoumbou Feunang et al. 2016. J. Cheminform.) via ClassyFire Batch by Fiehn Lab (https://cfb.fiehnlab.ucdavis.edu/) based on the PubChem InChIKey values.
5. Calculation of the detection frequency and of the mean and range of the retention time and similarity.
6. Writing of the result to an `Excel` file, plus a `.log` file with a full run report.

InChIKeys and classes are stored in cache files, so compounds that were already processed are not retrieved again in later runs.

## Prerequisites

Before using the script, the following applications have to be installed:

1. Python 3 (https://www.python.org/downloads/windows/). `tkinter` (the file dialogs) is included in the standard Windows installer.
2. Google Chrome. ChromeDriver does **not** have to be installed separately: Selenium downloads the version that matches the installed Chrome automatically. An outdated `chromedriver.exe` in the system `PATH` is ignored by the script.
3. The Python modules listed below.

Modules can be installed via Command Prompt (or the VS Code terminal) as follows:

* Type `cmd` in the Search line >> Click on the Command Prompt icon >> Type: `pip install pandas numpy openpyxl pubchempy selenium` >> Press `Enter`.

Module names:

* pandas
* numpy
* openpyxl
* pubchempy
* selenium

## How to use the script

The script can be run in VS Code or in Command Prompt:

* **VS Code:** open the script and click "Run Python File" (the triangle at the top right).
* **Command Prompt:** type `python "the path to the script including the extension (.py)"`, e.g. `python "D:/Projects/VOC_HTChar_Align_1.5.py"` >> Press `Enter`.
* Optional: the input (and output) file can be given directly, e.g. `python "D:/Projects/VOC_HTChar_Align_1.5.py" "D:/Projects/input.xlsx"`. In this case no file dialogs are shown.

Then:

1. Choose the file for processing in the pop-up window and press `Open`.
2. Choose where to save the output in the second pop-up window. Press `Cancel` to save it next to the input file as `<input name>_Aligned_&_ClassyFyed.xlsx`.
3. During the ClassyFire step, Chrome windows open and close automatically. Do not close them manually.

## Input file

The script is tailored to the output file format of the GCMSsolution (v4.52) software (Shimadzu, Japan).

* Every sheet is one sample, and the sheet name is used as the sample name.
* Each sheet must contain the following columns: `"Substance"`, `"RT, min"`, and `"Area"`. `"CAS"` and `"Similarity, %"` are optional (left empty if missing). Column `"#"` and any other columns are ignored.
* Small deviations in the column names are accepted (e.g. `"RT,min"`, `" Area "`, `"Similarity"`, `"Name"`), and the header row does not have to be the first row.
* Empty sheets, sheets without the required columns, and rows without a name or a valid retention time are skipped and reported in the log.
* CAS# values are normalized (e.g. `110-0-9` >> `110-00-9`, `0-0-0` >> `NA`). Decimal commas (e.g. `5,2`) are accepted.

## Output files

All output files are saved in the folder of the output file:

* **`<name>.xlsx`**: the aligned table, one row per compound and retention time group, sorted by mean retention time. Columns:
  * `Substance`, `CAS`, `InChIKey`, `InChIKey source` (`name`, `CAS`, `(cache)`, or `not found`)
  * `Number of detections`, `DF, %` (detection frequency), `RT_mean, min`, `RT_range, min`, `Similarity_mean, %`, `Similarity_range, %`
  * `RT, min_<sample>`, `Similarity, %_<sample>`, and `Area_<sample>` for each sample
  * `Class source` (`cache` or `ClassyFire`), `ClassyFire status`, `Kingdom`, `Superclass`, `Class`, `Subclass`, `Parent Level 1`–`5`
* **`<name>.log`**: run report with all warnings (skipped rows, duplicate peaks, compounds not found, ClassyFire problems).
* **`<name>_ClassyFire_debug_batch<N>.html/.png`**: only created if a ClassyFire batch gets stuck; a snapshot of the page for checking.

If the output file is open in Excel, the result is saved under a time-stamped name instead.

## Cache files

Two cache files are created next to the script and shared by all runs:

* **`PubChem_InChIKey_cache.csv`**: InChIKeys per compound name/CAS#, including compounds not found in PubChem.
* **`ClassyFire_class_cache.csv`**: ClassyFire results per InChIKey, including compounds without a class and compounds skipped because they blocked ClassyFire.

New results are written to the cache immediately, so nothing is lost if a run stops. Results that failed because of a server or connection error are not cached and are retried in the next run. To start fresh, delete the cache file(s).

If a cache file is open in Excel while the script runs, new entries are saved to `..._unsaved_<time>.csv` and merged into the cache automatically in the next run. Avoid saving the cache files from Excel, because Excel may convert CAS# values into dates.

## ClassyFire Batch

ClassyFire Batch sometimes gets stuck on a compound: the remaining compounds stay `Queued` and the page does not finish. The script handles this as follows:

* InChIKeys are submitted in batches (default: 200) in a fresh Chrome window per batch.
* If no progress is made for 5 min, the finished results are kept and cached, and the remaining InChIKeys are submitted again in a new batch.
* If a new batch makes no progress at all, its first InChIKey is considered to block ClassyFire. It is skipped and cached as `Skipped: ClassyFire stalled`.
* After 3 batches in a row without any progress, the ClassyFire step stops (the server is probably down), and the output is saved with the classes retrieved so far and the cached classes.

## Settings

All settings are at the top of the script (`SETTINGS` block). Nothing else has to be changed.

| Setting | Default | Description |
|---|---|---|
| `RT_TOLERANCE_MIN` | `10 / 60` | Retention time tolerance for alignment, min. |
| `DUPLICATE_PEAK_POLICY` | `'max_area'` | Same compound twice in one sample within the tolerance: keep the larger peak (`'max_area'`) or sum the areas (`'sum'`). |
| `RUN_PUBCHEM` | `True` | Retrieve InChIKeys. |
| `USE_CAS_FALLBACK` | `True` | Search PubChem by CAS# if the name is not found. |
| `PUBCHEM_PAUSE_S` | `0.25` | Pause between PubChem requests, s (PubChem allows max. 5 requests/s). |
| `PUBCHEM_RETRIES` | `3` | Retries if PubChem is busy or the connection fails. |
| `PUBCHEM_RETRY_NOT_FOUND` | `False` | Query compounds cached as "not found" again. |
| `RUN_CLASSYFIRE` | `True` | Retrieve new classes from ClassyFire Batch (`False` = cached classes only). |
| `CLASSYFIRE_BATCH_SIZE` | `200` | InChIKeys per ClassyFire batch. |
| `CLASSYFIRE_STALL_MIN` | `5` | Minutes without progress before a batch is considered stuck. |
| `CLASSYFIRE_POLL_S` | `15` | How often the progress is checked, s. |
| `CLASSYFIRE_MAX_FAILED_BATCHES` | `3` | Stop ClassyFire after this many batches in a row without progress. |
| `CLASSYFIRE_MAX_WAIT_S` | `2 * 3600` | Max. time per batch, s. |
| `CLASSYFIRE_INPUT_MODE` | `'one_by_one'` | How InChIKeys are entered: typed one by one (reliable) or `'javascript'` (pasted at once; faster but less tested). |
| `CLASSYFIRE_RETRY_UNCLASSIFIED` | `False` | Send cached InChIKeys without a class to ClassyFire again. |
| `CLASSYFIRE_RETRY_STALLED` | `False` | Send cached InChIKeys that blocked ClassyFire again. |

## Notes and recommendations

* Check the `.log` file after each run, especially the warnings.
* The alignment requires identical compound names and CAS# across samples. The same compound with different library names or CAS# values ends up in separate rows.
* Running all compounds through ClassyFire for the first time can take several hours. Later runs are much faster thanks to the cache.

## License

[![MIT License](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/license/mit)

Intended for academic and research use.
