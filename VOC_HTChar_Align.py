"""
A script used for inter-sample data alignment and classification
of Shimadzu TD-GC-MS result tables.

Version: 1.5 (2026-11-06)

Author: Andriy Rebryk, andrew.rebryk@gmail.com

What it does:
    1. Reads an Excel file where every sheet is one sample
       (columns: '#', 'Substance', 'CAS', 'Similarity, %', 'RT, min', 'Area').
    2. Cleans the tables (column names, whitespace, CAS format, numeric columns).
    3. Aligns the samples by Substance + CAS (exact match) and RT (within a tolerance).
    4. Retrieves InChIKeys from PubChem (by name; CAS as a fallback); all results are kept in a
       cache file (PubChem_InChIKey_cache.csv next to this script), so they are not retrieved again.
    5. Retrieves the chemical classes from ClassyFire Batch (https://cfb.fiehnlab.ucdavis.edu/); classified
       InChIKeys are kept in a cache file (ClassyFire_class_cache.csv next to this script).
       InChIKeys are sent in batches; if ClassyFire gets stuck on a compound, it is skipped & the rest continues.
    6. Calculates the detection frequency, mean & range of RT and Similarity.
    7. Writes the result to an Excel file (+ a .log file with the run report).

Requirements:
    Python 3, Google Chrome, and the modules below.
    Install them in the terminal (VS Code: Terminal --> New Terminal):
        pip install pandas numpy openpyxl pubchempy selenium
    (tkinter comes with the standard Python installer for Windows.)

How to run:
    - In VS Code: open this file and press "Run Python File" (the triangle top right), or
    - In the terminal: python "D:/Projects/VOC_HTChar_Align.py"
    - Optional: give the input file directly: python "VOC_HTChar_Align.py" "D:/Projects/input.xlsx"

    A window asks for the input Excel file, then a second window asks where to save the output
    (press "Cancel" to save it next to the input file with a default name).

All user settings are in the SETTINGS block below; nothing else has to be changed.
This script is tailored to a certain data table format.
Feel free to change the script or use any part of it for your needs.
"""

#importing modules
import logging
import os
import re
import sys
import time
from pathlib import Path

import numpy as np
import pandas as pd


#======================================================================================
# SETTINGS
#======================================================================================
RT_TOLERANCE_MIN = 10 / 60          #RT tolerance for alignment, min (10 s = ca. 0.1667 min)

#what to do if the same Substance + CAS occurs twice in ONE sample within the RT tolerance
#'max_area' --> keep the peak with the largest area (the other one is reported in the log)
#'sum'      --> sum the areas (RT & Similarity are taken from the largest peak)
DUPLICATE_PEAK_POLICY = 'sum'

RUN_PUBCHEM = True                  #retrieve InChIKeys from PubChem
USE_CAS_FALLBACK = True             #if the name is not found in PubChem, try the CAS number
PUBCHEM_PAUSE_S = 0.25              #pause between PubChem requests (PubChem allows max. 5 requests/s)
PUBCHEM_RETRIES = 3                 #number of retries if PubChem is busy / connection fails
#permanent InChIKey cache (CSV) next to this script; shared by all runs, so known compounds are not queried again
#you can open/edit it in Excel (close it before running the script) or delete it to start fresh
PUBCHEM_CACHE_FILE = Path(__file__).resolve().with_name('PubChem_InChIKey_cache.csv')
PUBCHEM_RETRY_NOT_FOUND = False     #True --> compounds cached as 'not found' are queried in PubChem again

RUN_CLASSYFIRE = True               #retrieve new classes from ClassyFire Batch via Chrome (False = cache only)
CLASSYFIRE_URL = 'https://cfb.fiehnlab.ucdavis.edu/'
CLASSYFIRE_BATCH_SIZE = 200         #InChIKeys per ClassyFire page; a stall then only affects one batch
CLASSYFIRE_STALL_MIN = 2            #no progress for this long, min --> batch is stalled; finished keys are kept &
                                    #the rest is sent again in a new Chrome window (a key that blocks is skipped)
CLASSYFIRE_POLL_S = 15              #how often the progress is checked, s
CLASSYFIRE_MAX_FAILED_BATCHES = 10  #stop ClassyFire after this many batches in a row without any progress (server down)
CLASSYFIRE_MAX_WAIT_S = 2 * 3600    #max. time per batch, s
#how InChIKeys are put into the ClassyFire text box:
#'one_by_one' --> typed one per line (reliable; ca. 1-2 min per 1000 keys)
#'javascript' --> pasted all at once (instant; try it if you want, switch back if ClassyFire ignores the input)
CLASSYFIRE_INPUT_MODE = 'one_by_one'
#permanent ClassyFire cache (CSV) next to this script; cached InChIKeys are not sent to ClassyFire again
#(delete the file to start fresh)
CLASSYFIRE_CACHE_FILE = Path(__file__).resolve().with_name('ClassyFire_class_cache.csv')
CLASSYFIRE_RETRY_UNCLASSIFIED = False  #True --> cached InChIKeys that ClassyFire finished WITHOUT a class are sent again
CLASSYFIRE_RETRY_STALLED = False       #True --> cached InChIKeys skipped because they blocked ClassyFire are sent again

DEFAULT_OUTPUT_SUFFIX = '_Aligned_&_ClassyFyed.xlsx'


#======================================================================================
# LOGGING SETUP (prints to the terminal and, later, to a .log file next to the output)
#======================================================================================
logging.basicConfig(
    level = logging.INFO,
    format = '%(asctime)s | %(levelname)-7s | %(message)s',
    datefmt = '%H:%M:%S',
    handlers = [logging.StreamHandler(sys.stdout)],
    force = True,
)
logging.getLogger('selenium').setLevel(logging.WARNING)  #silencing Selenium's own chatter
logging.getLogger('urllib3').setLevel(logging.WARNING)

pd.set_option('display.width', 250)
pd.set_option('display.max_columns', 30)

LINE = 150 * '*'


#======================================================================================
# COLUMN NAME HANDLING (foolproofing against renamed / misspelled headers)
#======================================================================================
#canonical column names used throughout the script
SUBSTANCE, CAS, SIM, RT, AREA = 'Substance', 'CAS', 'Similarity, %', 'RT, min', 'Area'

#header variants are reduced to lowercase letters/digits/% (e.g. ' RT (min) ' --> 'rtmin')
#and mapped to the canonical names
COLUMN_ALIASES = {
    'substance': SUBSTANCE, 'name': SUBSTANCE, 'compound': SUBSTANCE, 'compoundname': SUBSTANCE,
    'cas': CAS, 'casno': CAS, 'casnumber': CAS, 'casrn': CAS,
    'similarity%': SIM, 'similarity': SIM, 'si': SIM, 'si%': SIM, 'sim%': SIM,
    'rtmin': RT, 'rt': RT, 'retentiontime': RT, 'retentiontimemin': RT, 'rettime': RT,
    'area': AREA, 'peakarea': AREA,
}
REQUIRED_COLUMNS = [SUBSTANCE, RT, AREA]


def column_key(name):
    """Reducing a header to a comparable key: lowercase, letters/digits/% only."""
    return re.sub(r'[^a-z0-9%]', '', str(name).strip().lower())


def find_header_row(raw, max_rows = 15):
    """Finding the header row (the first row with a 'Substance'-like cell) in a sheet read with header = None."""
    for i in range(min(max_rows, len(raw))):
        keys = {column_key(v) for v in raw.iloc[i].tolist()}
        if any(COLUMN_ALIASES.get(k) == SUBSTANCE for k in keys):
            return i
    return None


#======================================================================================
# CLEANING FUNCTIONS
#======================================================================================
def clean_text(value):
    """Stripping & collapsing whitespace (incl. non-breaking spaces); NaN/None --> ''."""
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return ''
    text = str(value).replace('\u00a0', ' ')
    return re.sub(r'\s+', ' ', text).strip()


def correct_cas_number(cas):
    """Normalizing CAS numbers: '0-0-0'/empty --> 'NA'; '110-0-9' --> '110-00-9'."""
    cas = clean_text(cas).replace(' ', '')
    if cas == '' or cas.lower() in {'nan', 'na', 'n/a', 'none', '-'}:
        return 'NA'
    if re.fullmatch(r'0+-0+-0', cas):  #'0-0-0', '0-00-0' etc.
        return 'NA'
    #CAS stored as a number by Excel, e.g. 110009.0 --> '110-00-9'
    if re.fullmatch(r'\d+(\.0)?', cas):
        digits = cas.split('.')[0]
        if len(digits) >= 5:
            return f'{digits[:-3]}-{digits[-3:-1]}-{digits[-1]}'
        return cas
    parts = cas.split('-')
    if len(parts) == 3 and parts[0].isdigit() and parts[1].isdigit() and parts[2].isdigit():
        if len(parts[1]) == 1:  #correcting cases like '110-0-9' to '110-00-9'
            parts[1] = f'0{parts[1]}'
        return '-'.join(parts)
    #if no correction is possible, returning the CAS as is
    return cas


def cas_is_valid(cas):
    """Checking the CAS check digit (used to decide whether a CAS is worth a PubChem query)."""
    if not re.fullmatch(r'\d{2,7}-\d{2}-\d', str(cas)):
        return False
    digits = cas.replace('-', '')
    body, check = digits[:-1], int(digits[-1])
    total = sum(int(d) * (i + 1) for i, d in enumerate(reversed(body)))
    return total % 10 == check


def to_number(series, decimal_comma = True):
    """Converting a column to numbers; '12,5' --> 12.5 for RT/Similarity; non-numbers --> NaN."""
    if decimal_comma:
        series = series.map(lambda v: v.replace(',', '.').strip() if isinstance(v, str) else v)
    else:
        series = series.map(lambda v: re.sub(r'[\s\u00a0]', '', v) if isinstance(v, str) else v)
    return pd.to_numeric(series, errors = 'coerce')


def read_and_clean_sheets(xls_path):
    """Reading all sheets & returning one long, clean table with a 'Sample' column + the sample order."""
    raw_sheets = pd.read_excel(xls_path, header = None, sheet_name = None)  #reading the file once
    cleaned, samples = [], []

    for sheet_name, raw in raw_sheets.items():
        sample = clean_text(sheet_name)
        raw = raw.dropna(how = 'all').dropna(axis = 1, how = 'all').reset_index(drop = True)
        if raw.empty:
            logging.warning('Sheet "%s" is empty --> skipped.', sheet_name)
            continue

        header_row = find_header_row(raw)
        if header_row is None:
            logging.warning('Sheet "%s": no header row with a "Substance" column found --> skipped.', sheet_name)
            continue

        #setting the header row as column names & mapping header variants to canonical names
        headers = [COLUMN_ALIASES.get(column_key(h), clean_text(h)) for h in raw.iloc[header_row]]
        df = raw.iloc[header_row + 1:].copy()
        df.columns = headers

        if df.columns.duplicated().any():
            dupes = sorted(set(df.columns[df.columns.duplicated()]))
            logging.warning('Sheet "%s": duplicated columns %s --> only the first one is used.', sheet_name, dupes)
            df = df.loc[:, ~df.columns.duplicated()]

        missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
        if missing:
            logging.warning('Sheet "%s": missing required column(s) %s --> skipped. Found columns: %s',
                            sheet_name, missing, list(df.columns))
            continue
        if CAS not in df.columns:
            logging.warning('Sheet "%s": no "CAS" column --> CAS set to "NA".', sheet_name)
            df[CAS] = 'NA'
        if SIM not in df.columns:
            logging.warning('Sheet "%s": no "Similarity, %%" column --> left empty.', sheet_name)
            df[SIM] = np.nan

        ignored = [c for c in df.columns if c not in {SUBSTANCE, CAS, SIM, RT, AREA, '#', ''}]
        if ignored:
            logging.info('Sheet "%s": ignored column(s) %s', sheet_name, ignored)

        df = df[[SUBSTANCE, CAS, SIM, RT, AREA]].copy()
        df[SUBSTANCE] = df[SUBSTANCE].map(clean_text)
        df[CAS] = df[CAS].map(correct_cas_number)
        df[RT] = to_number(df[RT])
        df[SIM] = to_number(df[SIM])
        df[AREA] = to_number(df[AREA], decimal_comma = False)

        #dropping rows without a name or without a valid RT (they cannot be aligned)
        no_name = df[SUBSTANCE] == ''
        no_rt = df[RT].isna() & ~no_name
        if no_rt.any():
            logging.warning('Sheet "%s": %d row(s) without a valid RT --> dropped:\n%s',
                            sheet_name, no_rt.sum(), df[no_rt].to_string())
        df = df[~no_name & ~no_rt]

        if df.empty:
            logging.warning('Sheet "%s": no valid rows left --> skipped.', sheet_name)
            continue

        df.insert(0, 'Sample', sample)
        cleaned.append(df)
        samples.append(sample)
        logging.info('Input sheet "%s": %d peaks\n%s', sheet_name, len(df), df.head(10).to_string())
        logging.info(LINE)

    if not cleaned:
        raise ValueError('No usable sheets found in the input file. Check the column names.')

    long_df = pd.concat(cleaned, ignore_index = True)
    return long_df, samples


#======================================================================================
# ALIGNMENT
#======================================================================================
def align_samples(long_df, samples, tolerance = RT_TOLERANCE_MIN):
    """
    Aligning peaks across samples: exact match for Substance (case-insensitive) & CAS,
    approximate match for RT. Peaks are sorted by RT; a new group (= output row) starts when
    a peak is more than 'tolerance' away from the first (earliest) peak of the current group.
    """
    long_df = long_df.copy()
    long_df['_key'] = long_df[SUBSTANCE].str.casefold()
    aligned_rows = []
    duplicate_log = []

    for (_, cas), grp in long_df.groupby(['_key', CAS], sort = False):
        substance = grp[SUBSTANCE].iloc[0]  #keeping the first spelling of the name
        grp = grp.sort_values(RT)

        #grouping RT values within the tolerance range
        groups, current, group_start = [], [], None
        for _, peak in grp.iterrows():
            if group_start is None or abs(peak[RT] - group_start) <= tolerance:
                current.append(peak)
                if group_start is None:
                    group_start = peak[RT]
            else:
                groups.append(current)
                current, group_start = [peak], peak[RT]
        groups.append(current)

        #creating one output row per group
        for group in groups:
            row = {SUBSTANCE: substance, CAS: cas}
            for peak in group:
                s = peak['Sample']
                area_col = f'{AREA}_{s}'
                if area_col in row:  #the same sample already has a peak in this group
                    old_area = row[area_col]
                    new_area = peak[AREA]
                    duplicate_log.append(f'{substance} ({cas}) in "{s}": RT {row[f"{RT}_{s}"]:.3f} & {peak[RT]:.3f}')
                    bigger_is_new = (pd.notna(new_area) and (pd.isna(old_area) or new_area > old_area))
                    if DUPLICATE_PEAK_POLICY == 'sum':
                        summed = np.nansum([old_area, new_area])
                        if bigger_is_new:
                            row[f'{RT}_{s}'], row[f'{SIM}_{s}'] = peak[RT], peak[SIM]
                        row[area_col] = summed
                        continue
                    if not bigger_is_new:  #'max_area': keeping the old (larger) peak
                        continue
                row[f'{RT}_{s}'] = peak[RT]
                row[f'{SIM}_{s}'] = peak[SIM]
                row[area_col] = peak[AREA]
            aligned_rows.append(row)

    if duplicate_log:
        logging.warning('%d case(s) of the same compound twice in one sample within the RT tolerance '
                        '(policy: %s):\n  %s', len(duplicate_log), DUPLICATE_PEAK_POLICY, '\n  '.join(duplicate_log))

    #fixed column order: Substance, CAS, then RT / Similarity / Area for each sample in sheet order
    sample_cols = [f'{c}_{s}' for s in samples for c in (RT, SIM, AREA)]
    df_aligned = pd.DataFrame(aligned_rows).reindex(columns = [SUBSTANCE, CAS] + sample_cols)
    return df_aligned


#======================================================================================
# PUBCHEM (InChIKeys) with a permanent cache file
#======================================================================================
INCHIKEY_PATTERN = re.compile(r'^[A-Z]{14}-[A-Z]{10}-[A-Z]$')
CACHE_COLUMNS = ['Query', 'Namespace', 'InChIKey', 'Retrieved']


def cache_key(query, namespace):
    """Cache key: namespace + cleaned, case-insensitive query (e.g. ('name', 'toluene'))."""
    return (namespace, clean_text(query).casefold())


def merge_unsaved_cache_files(cache_path, columns):
    """Adding entries from '..._unsaved_*.csv' files (left by a run while the cache was open in Excel) to the cache."""
    unsaved = sorted(cache_path.parent.glob(f'{cache_path.stem}_unsaved_*.csv'))
    if not unsaved:
        return
    try:
        rows = pd.concat([pd.read_csv(f, dtype = str, keep_default_na = False, encoding = 'utf-8-sig') for f in unsaved])
        new_file = not cache_path.is_file()
        rows = rows.reindex(columns = columns, fill_value = '')
        rows.to_csv(cache_path, mode = 'a', header = new_file, index = False,
                                   encoding = 'utf-8-sig' if new_file else 'utf-8')
    except Exception as err:
        logging.warning('Could not merge unsaved cache file(s) %s into the cache (%s) --> will retry next run.',
                        [f.name for f in unsaved], err)
        return
    for f in unsaved:
        f.unlink()
    logging.info('Merged %d entries from %d unsaved cache file(s) into the cache file.', len(rows), len(unsaved))


def load_inchikey_cache(cache_path):
    """Reading the cache file into a dictionary {(namespace, query): InChIKey}; '' = not found in PubChem."""
    merge_unsaved_cache_files(cache_path, CACHE_COLUMNS)
    cache = {}
    if not cache_path.is_file():
        logging.info('No InChIKey cache file yet --> a new one will be created: %s', cache_path)
        return cache
    try:
        df = pd.read_csv(cache_path, dtype = str, keep_default_na = False, encoding = 'utf-8-sig')
    except Exception as err:
        logging.warning('InChIKey cache file cannot be read (%s) --> starting with an empty cache.', err)
        return cache
    bad = 0
    for q, ns, key in zip(df.get('Query', []), df.get('Namespace', []), df.get('InChIKey', [])):
        key = clean_text(key).upper()
        if key and not INCHIKEY_PATTERN.match(key):  #skipping damaged / hand-edited wrong entries
            bad += 1
            continue
        cache[cache_key(q, ns)] = key  #later lines overwrite earlier ones (newest result wins)
    if bad:
        logging.warning('%d invalid InChIKey(s) in the cache file were ignored.', bad)
    logging.info('InChIKey cache loaded: %d entries (%d with InChIKey, %d "not found") from %s',
                 len(cache), sum(1 for v in cache.values() if v), sum(1 for v in cache.values() if not v), cache_path)
    return cache


class CacheWriter:
    """Appending new results to a cache file immediately (nothing is lost if the run stops); used by both caches."""
    def __init__(self, cache_path, columns):
        self.path = cache_path
        self.columns = columns
        self.failed = []  #rows that could not be written (e.g. the file is open in Excel)

    def add(self, rows):
        """rows = list of lists in the order of 'columns'."""
        if not rows:
            return
        self.failed.append(pd.DataFrame(rows, columns = self.columns))
        try:
            new_file = not self.path.is_file()
            pd.concat(self.failed).to_csv(self.path, mode = 'a', header = new_file, index = False,
                                          encoding = 'utf-8-sig' if new_file else 'utf-8')
            self.failed = []
        except PermissionError:
            if len(self.failed) == 1:
                logging.warning('Cache file "%s" is open in another program (Excel?) --> '
                                'new entries are kept and written when it is closed.', self.path.name)

    def close(self):
        if self.failed:
            alt = self.path.with_name(f'{self.path.stem}_unsaved_{time.strftime("%Y%m%d_%H%M%S")}.csv')
            pd.concat(self.failed).to_csv(alt, index = False, encoding = 'utf-8-sig')
            logging.warning('%d new cache entries could not be added to "%s" --> saved to "%s"; '
                            'they are merged into the cache automatically next run.', len(self.failed), self.path.name, alt.name)


def pubchem_inchikey(query, namespace = 'name'):
    """
    Returning the InChIKey of the first PubChem hit for a name (or CAS):
    'XXXX...' = found, '' = not found, None = PubChem error (result is NOT cached, so it is retried next run).
    """
    import pubchempy as pcp
    for attempt in range(1, PUBCHEM_RETRIES + 1):
        try:
            time.sleep(PUBCHEM_PAUSE_S)
            results = pcp.get_compounds(query, namespace)
            for compound in results:
                if compound.inchikey:
                    return compound.inchikey
            return ''
        except Exception as err:  #'not found' is returned as an empty list, so errors here are server/connection issues
            if 'NotFound' in type(err).__name__:
                return ''
            logging.warning('PubChem error for "%s" (attempt %d/%d): %s', query, attempt, PUBCHEM_RETRIES, err)
            time.sleep(2 * attempt)
    return None


def add_inchikeys(df):
    """
    Adding 'InChIKey' & 'InChIKey source' columns.
    Every unique name/CAS is looked up only once: first in the cache file, then in PubChem.
    """
    df = df.copy()
    cache = load_inchikey_cache(PUBCHEM_CACHE_FILE)
    writer = CacheWriter(PUBCHEM_CACHE_FILE, CACHE_COLUMNS)
    counts = {'cache': 0, 'PubChem': 0, 'error': 0}

    def lookup(query, namespace):
        """Returning (InChIKey, from_cache); InChIKey = '' if not found or on error."""
        k = cache_key(query, namespace)
        if k in cache and (cache[k] or not PUBCHEM_RETRY_NOT_FOUND):
            counts['cache'] += 1
            return cache[k], True
        key = pubchem_inchikey(query, 'name')  #CAS numbers are searched as PubChem synonyms ('name')
        if key is None:  #PubChem error --> not cached
            counts['error'] += 1
            return '', False
        counts['PubChem'] += 1
        cache[k] = key
        writer.add([[clean_text(query), namespace, key, time.strftime('%Y-%m-%d')]])
        return key, False

    pairs = list(df[[SUBSTANCE, CAS]].drop_duplicates().itertuples(index = False, name = None))
    results = {}
    logging.info('Retrieving InChIKeys for %d unique compounds (cache first, then PubChem). It may take a while...',
                 len(pairs))
    try:
        for n, (name, cas) in enumerate(pairs, start = 1):
            key, cached = lookup(name, 'name')
            source = 'name'
            if not key and USE_CAS_FALLBACK and cas_is_valid(cas):
                key, cached = lookup(cas, 'CAS')
                source = 'CAS'
            if not key:
                source, cached = 'not found', False
            elif cached:
                source += ' (cache)'
            results[(name, cas)] = (key, source)
            logging.info('%d/%d  %s (%s) --> %s [%s]', n, len(pairs), name, cas, key or '-', source)
    finally:
        writer.close()  #saving any cache entries that could not be written yet

    keys = [results[(n, c)] for n, c in zip(df[SUBSTANCE], df[CAS])]
    df['InChIKey'] = [k for k, _ in keys]
    df['InChIKey source'] = [s for _, s in keys]

    logging.info('InChIKey lookups: %d from cache, %d from PubChem, %d PubChem errors (will be retried next run).',
                 counts['cache'], counts['PubChem'], counts['error'])
    not_found = sorted({n for n, c in pairs if results[(n, c)][1] == 'not found'})
    if not_found:
        logging.warning('%d compound(s) without InChIKey:\n  %s', len(not_found), '\n  '.join(not_found))
    return df


#======================================================================================
# CLASSYFIRE (classes)
#======================================================================================
CLASS_COLUMNS = ['InChIKey', 'Status', 'Kingdom', 'Superclass', 'Class', 'Subclass',
                 'Parent Level 1', 'Parent Level 2', 'Parent Level 3', 'Parent Level 4', 'Parent Level 5']
CLASS_CACHE_COLUMNS = CLASS_COLUMNS + ['Retrieved']
INCHIKEY_IN_TEXT = re.compile(r'[A-Z]{14}-[A-Z]{10}-[A-Z]')
STALLED_STATUS = 'Skipped: ClassyFire stalled'

#JavaScript that finds & reads the whole result table in ONE call (reading cell by cell = thousands of slow calls)
JS_READ_TABLE = """
const t = document.querySelector('#results table') || document.querySelector('table.table') || document.querySelector('.table');
if (!t) { return []; }
return Array.from(t.querySelectorAll('tr')).map(
    tr => Array.from(tr.querySelectorAll('td, th')).map(c => (c.innerText || c.textContent || '').trim()));
"""

#JavaScript that pastes all InChIKeys into the text box at once (only used if CLASSYFIRE_INPUT_MODE = 'javascript')
JS_SET_TEXT = """
arguments[0].value = arguments[1];
arguments[0].dispatchEvent(new Event('input', {bubbles: true}));
arguments[0].dispatchEvent(new Event('change', {bubbles: true}));
"""


def parse_classyfire_rows(rows):
    """
    Turning raw table rows into [InChIKey, Status, Kingdom, ...] lists.
    The InChIKey is searched in every cell (not assumed to be the 1st cell); rows without one (header) are skipped.
    """
    data = []
    for cells in rows:
        for i, cell in enumerate(cells):
            m = INCHIKEY_IN_TEXT.search(str(cell).upper())
            if m:
                rest = [str(c).strip() for c in cells[i + 1:]]
                data.append(([m.group(0)] + rest + [''] * len(CLASS_COLUMNS))[:len(CLASS_COLUMNS)])
                break
    return data


def is_pending(status):
    """True if ClassyFire has not finished this InChIKey yet ('Queued', 'Processing', ... or no status)."""
    s = clean_text(status).lower()
    return s == '' or any(w in s for w in ('queue', 'process', 'pend', 'running', 'wait', 'submit'))


def save_debug_page(driver, debug_path, reason):
    """Saving the ClassyFire page (HTML + screenshot) so the problem can be checked afterwards."""
    if driver is None or debug_path is None:
        return
    try:
        debug_path.write_text(driver.page_source, encoding = 'utf-8')
        driver.save_screenshot(str(debug_path.with_suffix('.png')))
        logging.warning('%s --> ClassyFire page saved for checking: %s (+ .png)', reason, debug_path.name)
    except Exception as err:
        logging.warning('%s --> ClassyFire page could not be saved (%s).', reason, type(err).__name__)


def start_chrome():
    """Launching Chrome; old chromedriver.exe files in PATH are hidden, so Selenium fetches the matching driver."""
    from selenium import webdriver
    path_dirs = os.environ.get('PATH', '').split(os.pathsep)
    driver_dirs = [d for d in path_dirs if d and (Path(d) / 'chromedriver.exe').is_file()]
    if driver_dirs:
        logging.info('Ignoring chromedriver in PATH (Selenium will fetch the matching one): %s', driver_dirs)
        os.environ['PATH'] = os.pathsep.join(d for d in path_dirs if d not in driver_dirs)
    try:
        driver = webdriver.Chrome()
    finally:
        os.environ['PATH'] = os.pathsep.join(path_dirs)  #restoring PATH
    driver.set_page_load_timeout(120)
    return driver


def run_classyfire_page(batch, batch_no, debug_path = None):
    """
    Submitting one batch of InChIKeys in a fresh Chrome window & watching the progress.
    Returns (rows, outcome): rows = [InChIKey, Status, Kingdom, ...] for the batch as last read from the page;
    outcome = 'finished' / 'stalled' (no progress for CLASSYFIRE_STALL_MIN) / 'not responding' / 'timeout'.
    Errors before the batch is submitted (e.g. Chrome cannot start) are raised.
    """
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    batch_set = set(batch)
    rows, outcome = [], None
    driver = start_chrome()
    try:
        driver.get(CLASSYFIRE_URL)
        textarea = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'inchikeys')))
        button = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='submit']")))

        #inserting InChIKeys
        textarea.clear()
        if CLASSYFIRE_INPUT_MODE == 'javascript':
            driver.execute_script(JS_SET_TEXT, textarea, '\n'.join(batch) + '\n')
            logging.info('Batch %d: inserted %d InChIKeys at once (JavaScript).', batch_no, len(batch))
        else:
            #one by one: many short commands instead of one huge one, which can freeze Chrome / time out
            for n, key in enumerate(batch, start = 1):
                textarea.send_keys(key + '\n')
                if n % 100 == 0 or n == len(batch):
                    logging.info('Batch %d: inserted %d/%d InChIKeys', batch_no, n, len(batch))

        #clicking the ClassyFy button after entering all InChIKeys
        button.click()
        logging.info('Batch %d submitted; checking progress every %d s (stall limit: %.0f min)...',
                     batch_no, CLASSYFIRE_POLL_S, CLASSYFIRE_STALL_MIN)

        #watching the progress: finished = no 'Queued'/'Processing' rows left
        start = last_change = time.time()
        last_done = -1
        while outcome is None:
            time.sleep(CLASSYFIRE_POLL_S)
            try:
                raw = driver.execute_script(JS_READ_TABLE) or []
            except Exception as err:  #a frozen page does not answer
                logging.warning('Batch %d: ClassyFire page not responding (%s).', batch_no, type(err).__name__)
                outcome = 'not responding'
                break
            parsed = [r for r in parse_classyfire_rows(raw) if r[0] in batch_set]
            if parsed:
                rows = parsed  #keeping the latest successful read
            n_done = sum(not is_pending(r[1]) for r in rows)
            if n_done != last_done:
                logging.info('Batch %d: %d/%d done', batch_no, n_done, len(batch))
                last_done, last_change = n_done, time.time()
            if rows and n_done == len(rows):
                outcome = 'finished'
            elif time.time() - last_change > CLASSYFIRE_STALL_MIN * 60:
                stuck = next((r[0] for r in rows if is_pending(r[1])), '?')
                logging.warning('Batch %d: no progress for %.0f min (stuck before/at %s).',
                                batch_no, CLASSYFIRE_STALL_MIN, stuck)
                outcome = 'stalled'
            elif time.time() - start > CLASSYFIRE_MAX_WAIT_S:
                logging.warning('Batch %d: max. waiting time reached.', batch_no)
                outcome = 'timeout'

        if outcome in ('stalled', 'timeout') and debug_path is not None:
            save_debug_page(driver, debug_path.with_name(f'{debug_path.stem}_batch{batch_no}.html'),
                            f'Batch {batch_no} {outcome}')
    finally:
        try:
            driver.quit()  #closing the browser (also a frozen one)
        except Exception:
            pass
    return rows, outcome


def classyfire_batch(inchikeys, on_result, debug_path = None):
    """
    Getting classes from ClassyFire Batch in batches of CLASSYFIRE_BATCH_SIZE InChIKeys.
    If a batch stalls, the finished InChIKeys are kept & the rest is sent again in a new Chrome window.
    If a new batch makes no progress at all, its first InChIKey is the culprit: it is skipped
    (& cached as 'Skipped: ClassyFire stalled'). Results are passed to on_result() after every batch.
    """
    logging.info('Retrieving classes from ClassyFire for %d unique InChIKeys (batches of %d). It will take a while...',
                 len(inchikeys), CLASSYFIRE_BATCH_SIZE)
    remaining = list(inchikeys)
    done_total, no_progress_streak, batch_no = 0, 0, 0
    pending_skips = []  #skipped culprits; cached only once ClassyFire is shown to work (not if the server is down)

    while remaining:
        batch_no += 1
        batch = remaining[:CLASSYFIRE_BATCH_SIZE]
        logging.info('Batch %d: %d InChIKeys (%d of %d still to do)', batch_no, len(batch), len(remaining), len(inchikeys))
        rows, outcome = run_classyfire_page(batch, batch_no, debug_path)

        done = {r[0]: r for r in rows if not is_pending(r[1])}  #one row per finished InChIKey
        if done:
            on_result(list(done.values()))
            done_total += len(done)
            no_progress_streak = 0
            if pending_skips:
                on_result(pending_skips)
                pending_skips = []
        remaining = [k for k in remaining if k not in done]

        if outcome == 'finished':
            not_returned = [k for k in batch if k not in done]
            if not_returned:
                logging.warning('Batch %d: %d InChIKey(s) not returned by ClassyFire (retried next run):\n  %s',
                                batch_no, len(not_returned), '\n  '.join(not_returned))
                remaining = [k for k in remaining if k not in set(not_returned)]
        elif done:
            logging.warning('Batch %d %s after %d/%d InChIKeys --> the rest is sent again in a new batch.',
                            batch_no, outcome, len(done), len(batch))
        else:
            no_progress_streak += 1
            if no_progress_streak >= CLASSYFIRE_MAX_FAILED_BATCHES:
                raise RuntimeError(f'ClassyFire made no progress in {no_progress_streak} batches in a row '
                                   f'--> the server is probably down or overloaded; try again later.')
            culprit = batch[0]
            logging.warning('Batch %d: no InChIKey finished --> %s probably blocks ClassyFire & is skipped.',
                            batch_no, culprit)
            pending_skips.append([culprit, STALLED_STATUS] + [''] * (len(CLASS_COLUMNS) - 2))
            remaining.remove(culprit)

    if pending_skips and done_total:
        on_result(pending_skips)
    logging.info('ClassyFire finished: %d InChIKeys done in %d batch(es).', done_total, batch_no)


def load_classyfire_cache(cache_path):
    """Reading the ClassyFire cache file into a table with one row per InChIKey (newest entry wins)."""
    merge_unsaved_cache_files(cache_path, CLASS_CACHE_COLUMNS)
    empty = pd.DataFrame(columns = CLASS_COLUMNS)
    if not cache_path.is_file():
        logging.info('No ClassyFire cache file yet --> a new one will be created: %s', cache_path)
        return empty
    try:
        df = pd.read_csv(cache_path, dtype = str, keep_default_na = False, encoding = 'utf-8-sig')
    except Exception as err:
        logging.warning('ClassyFire cache file cannot be read (%s) --> starting with an empty cache.', err)
        return empty
    df = df.reindex(columns = CLASS_COLUMNS, fill_value = '')
    df['InChIKey'] = df['InChIKey'].map(clean_text).str.upper()
    valid = df['InChIKey'].str.match(INCHIKEY_PATTERN)
    if (~valid).any():
        logging.warning('%d row(s) with an invalid InChIKey in the ClassyFire cache file were ignored.', (~valid).sum())
    df = df[valid].drop_duplicates('InChIKey', keep = 'last')
    logging.info('ClassyFire cache loaded: %d InChIKeys from %s', len(df), cache_path)
    return df


def get_classes(inchikeys, debug_path = None):
    """
    Getting classes for unique InChIKeys: first from the ClassyFire cache file, the rest from ClassyFire Batch.
    Every finished InChIKey is cached right after its batch (classified, without class, or skipped as stalled).
    Returns a table with one row per InChIKey + 'Class source' ('cache' / 'ClassyFire') & 'ClassyFire status'.
    """
    cache = load_classyfire_cache(CLASSYFIRE_CACHE_FILE)
    cache = cache[cache['InChIKey'].isin(inchikeys)].copy()
    has_class = cache['Kingdom'].map(clean_text) != ''
    stalled = cache['Status'].str.startswith('Skipped')
    no_class = ~has_class & ~stalled
    keep = has_class | (stalled & (not CLASSYFIRE_RETRY_STALLED)) | (no_class & (not CLASSYFIRE_RETRY_UNCLASSIFIED))
    cached = cache[keep].copy()
    cached['Class source'] = 'cache'
    cached_keys = set(cached['InChIKey'])
    to_query = [k for k in inchikeys if k not in cached_keys]
    logging.info('Classes from cache: %d classified, %d without class, %d skipped as stalled; %d to send to ClassyFire.',
                 (has_class & keep).sum(), (no_class & keep).sum(), (stalled & keep).sum(), len(to_query))

    new_rows = []
    if to_query and RUN_CLASSYFIRE:
        writer = CacheWriter(CLASSYFIRE_CACHE_FILE, CLASS_CACHE_COLUMNS)

        def on_result(rows):  #caching every batch right away, so nothing is lost if a later batch fails
            new_rows.extend(rows)
            writer.add([r + [time.strftime('%Y-%m-%d')] for r in rows])

        try:
            classyfire_batch(to_query, on_result, debug_path = debug_path)
        except Exception as err:  #classes retrieved so far & cached classes are still used
            msg = getattr(err, 'msg', None) or str(err) or repr(err)  #Selenium errors keep their text in .msg
            logging.error('ClassyFire step stopped --> classes retrieved so far & cached classes are used.\n%s: %s',
                          type(err).__name__, msg)
        finally:
            writer.close()
    elif to_query:
        logging.info('ClassyFire Batch skipped (RUN_CLASSYFIRE = False) --> only cached classes are used.')

    new = pd.DataFrame(new_rows, columns = CLASS_COLUMNS)
    if not new.empty:
        new_class = (new['Kingdom'].map(clean_text) != '').sum()
        new_stalled = new['Status'].str.startswith('Skipped').sum()
        logging.info('From ClassyFire this run: %d classified, %d without class, %d skipped as stalled.',
                     new_class, len(new) - new_class - new_stalled, new_stalled)
    new['Class source'] = 'ClassyFire'

    out = pd.concat([cached, new], ignore_index = True).drop_duplicates('InChIKey', keep = 'last')
    return out.rename(columns = {'Status': 'ClassyFire status'})


#======================================================================================
# STATISTICS & OUTPUT
#======================================================================================
def range_text(row, fmt):
    """'min - max' as text; '' if the row has no values."""
    values = row.dropna()
    if values.empty:
        return ''
    return f'{values.min():{fmt}} - {values.max():{fmt}}'


def add_statistics(df, samples):
    rt_cols = [f'{RT}_{s}' for s in samples]
    sim_cols = [f'{SIM}_{s}' for s in samples]
    area_cols = [f'{AREA}_{s}' for s in samples]
    sample_cols = rt_cols + sim_cols + area_cols

    #converting the per-sample columns to numbers in one go (a new block, not column by column)
    num = df[sample_cols].apply(pd.to_numeric, errors = 'coerce')

    #a detection = the sample has a peak in this row (RT is always present for an aligned peak)
    n_det = num[rt_cols].notna().sum(axis = 1)
    stats = pd.DataFrame({
        'Number of detections': n_det,
        'DF, %': (n_det / len(samples) * 100).round(0).astype(int),
        'RT_mean, min': num[rt_cols].mean(axis = 1).round(3),
        'RT_range, min': num[rt_cols].apply(range_text, axis = 1, fmt = '.3f'),
        'Similarity_mean, %': num[sim_cols].mean(axis = 1).round(1),
        'Similarity_range, %': num[sim_cols].apply(range_text, axis = 1, fmt = '.0f'),
    }, index = df.index)

    #joining all columns at once (avoids the 'DataFrame is highly fragmented' PerformanceWarning)
    return pd.concat([df.drop(columns = sample_cols), num, stats], axis = 1)


def save_excel(df, out_path):
    """Writing the output; if the file is open in Excel, saving under a time-stamped name instead."""
    try:
        df.to_excel(out_path, sheet_name = 'OUT', index = False)
        return out_path
    except PermissionError:
        alt = out_path.with_name(f'{out_path.stem}_{time.strftime("%Y%m%d_%H%M%S")}{out_path.suffix}')
        logging.warning('"%s" is open in another program (Excel?) --> saving as "%s".', out_path.name, alt.name)
        df.to_excel(alt, sheet_name = 'OUT', index = False)
        return alt


def ask_paths():
    """Getting input & output paths (command-line argument or file dialogs)."""
    if len(sys.argv) > 1:
        in_path = Path(sys.argv[1])
        out_path = Path(sys.argv[2]) if len(sys.argv) > 2 else None
    else:
        import tkinter as tk
        from tkinter.filedialog import askopenfilename, asksaveasfilename
        root = tk.Tk()
        root.withdraw()                   #preventing the empty tkinter window from coming up
        root.attributes('-topmost', True)  #bringing the dialog to the front (it can hide behind VS Code)
        in_path = askopenfilename(title = 'Select Excel file for processing...',
                                  filetypes = [('Excel files', '*.xlsx *.xls *.xlsm'), ('All files', '*.*')],
                                  parent = root)
        if not in_path:
            root.destroy()
            sys.exit('No input file selected --> stopped.')
        in_path = Path(in_path)
        out_path = asksaveasfilename(title = 'Save output as... (Cancel = next to the input file)',
                                     initialdir = str(in_path.parent),
                                     initialfile = in_path.stem + DEFAULT_OUTPUT_SUFFIX,
                                     defaultextension = '.xlsx', filetypes = [('Excel files', '*.xlsx')],
                                     parent = root)
        root.destroy()
        out_path = Path(out_path) if out_path else None

    if not in_path.is_file():
        sys.exit(f'Input file not found: {in_path}')
    if out_path is None:
        out_path = in_path.with_name(in_path.stem + DEFAULT_OUTPUT_SUFFIX)
    if out_path.suffix.lower() != '.xlsx':
        out_path = out_path.with_suffix('.xlsx')
    if out_path.resolve() == in_path.resolve():
        sys.exit('The output file must not overwrite the input file.')
    return in_path, out_path


#======================================================================================
# MAIN
#======================================================================================
def main():
    in_path, out_path = ask_paths()

    #writing the log also to a file next to the output
    log_path = out_path.with_suffix('.log')
    file_handler = logging.FileHandler(log_path, mode = 'w', encoding = 'utf-8')
    file_handler.setFormatter(logging.Formatter('%(asctime)s | %(levelname)-7s | %(message)s', '%H:%M:%S'))
    logging.getLogger().addHandler(file_handler)

    start_time = time.strftime('%Y-%m-%d %H:%M:%S')
    logging.info(LINE)
    logging.info('Processing start time: %s', start_time)
    logging.info('Input file:  %s', in_path)
    logging.info('Output file: %s', out_path)
    logging.info('Cache file:  %s', PUBCHEM_CACHE_FILE)
    logging.info(LINE)

    #reading & cleaning
    logging.info('STEP 1/6: reading & cleaning the input sheets')
    logging.info(LINE)
    long_df, samples = read_and_clean_sheets(in_path)
    logging.info('Samples (%d): %s', len(samples), samples)
    logging.info(LINE)

    #aligning
    logging.info('STEP 2/6: aligning the samples (RT tolerance: %.4f min)', RT_TOLERANCE_MIN)
    logging.info(LINE)
    df = align_samples(long_df, samples, tolerance = RT_TOLERANCE_MIN)
    logging.info('Aligned table: %d rows from %d peaks\n%s', len(df), len(long_df), df.head(20).to_string())
    logging.info(LINE)

    #InChIKeys
    logging.info('STEP 3/6: InChIKeys')
    logging.info(LINE)
    if RUN_PUBCHEM:
        df = add_inchikeys(df)
    else:
        logging.info('Skipped (RUN_PUBCHEM = False).')
        df['InChIKey'], df['InChIKey source'] = '', 'skipped'
    logging.info(LINE)

    #classes; merged by InChIKey (not by row position), so the classes cannot shift to wrong compounds
    logging.info('STEP 4/6: ClassyFire classes')
    logging.info(LINE)
    class_cols = ['Class source', 'ClassyFire status'] + CLASS_COLUMNS[2:]
    unique_keys = sorted(k for k in df['InChIKey'].unique() if k)
    if not unique_keys:
        logging.warning('No InChIKeys available --> no classes.')
        df = df.assign(**{c: '' for c in class_cols})  #adding all empty class columns at once
    else:
        df_classes = get_classes(unique_keys, debug_path = out_path.with_name(out_path.stem + '_ClassyFire_debug.html'))
        df = df.merge(df_classes[['InChIKey'] + class_cols], on = 'InChIKey', how = 'left')
        n_with_key = (df['InChIKey'] != '').sum()
        n_classified = df['Kingdom'].fillna('').map(clean_text).ne('').sum()
        logging.info('Rows with a class: %d of %d rows with an InChIKey.', n_classified, n_with_key)
        if n_classified == 0:
            logging.error('No classes in the output. Check the ClassyFire messages above and the debug files.')
    logging.info(LINE)

    #statistics
    logging.info('STEP 5/6: statistics (detection frequency, RT & Similarity mean/range)')
    logging.info(LINE)
    df = add_statistics(df, samples)

    #column order: identifiers, statistics, per-sample data, classes
    stat_cols = ['Number of detections', 'DF, %', 'RT_mean, min', 'RT_range, min',
                 'Similarity_mean, %', 'Similarity_range, %']
    sample_cols = [f'{c}_{s}' for s in samples for c in (RT, SIM, AREA)]
    df = df[[SUBSTANCE, CAS, 'InChIKey', 'InChIKey source'] + stat_cols + sample_cols + class_cols]

    #sorting by mean RT & Substance
    df = df.sort_values(by = ['RT_mean, min', SUBSTANCE], ascending = True).reset_index(drop = True)
    logging.info('Out table:\n%s', df.head(20).to_string())
    logging.info(LINE)

    #writing to Excel (empty cells for missing values)
    logging.info('STEP 6/6: saving')
    logging.info(LINE)
    saved = save_excel(df, out_path)
    logging.info('Saved: %s (%d rows, %d samples)', saved, len(df), len(samples))
    logging.info('Log:   %s', log_path)
    logging.info(LINE)
    logging.info('Processing start time: %s', start_time)
    logging.info('Processing end time:   %s', time.strftime('%Y-%m-%d %H:%M:%S'))
    logging.info(LINE)
    logging.info(LINE)


if __name__ == '__main__':
    main()
