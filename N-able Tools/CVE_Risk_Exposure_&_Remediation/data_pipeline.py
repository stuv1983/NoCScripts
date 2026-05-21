"""
data_pipeline.py — all pandas data loading, merging, patch matching, and trend comparison.
No GUI imports. No xlsxwriter. Pure data in, data out.
"""

import logging
import os
import re
from datetime import datetime
from typing import Optional, Set, Tuple

import pandas as pd

from config import (
    CVE_PATTERN, PRODUCT_MAP, FIXED_VERSION_RULES,
    STATUS_RANK, STATUS_LABEL, INSTALLED_STATUSES,
    _CONFIG,
)
# Set up logging for the module. This will allow us to log debug information, warnings, and errors as needed throughout the data processing pipeline. By using a logger, we can provide insights into the internal workings of the code, which can be helpful for troubleshooting and understanding how the data is being processed at each step.   
# The logger is configured at the module level, so it can be used by all functions within this file to log relevant information. We can log things like the number of records loaded, any issues encountered during data processing, and summaries of the results of various operations. This will help us to maintain visibility into the data pipeline and ensure that we can identify and address any problems that arise.
# By using logging instead of print statements, we can also control the level of detail that is output and easily disable or redirect logs as needed without changing the code. This makes our data pipeline more robust and easier to maintain in the long run.
# Overall, setting up logging is an important step in building a reliable and maintainable data processing pipeline, and it will help us to ensure that we can effectively monitor and troubleshoot the code as we develop and use it.
log = logging.getLogger(__name__)

# python-calamine reads xlsx 4-5× faster than openpyxl via a Rust parser.
try:
    import python_calamine  # noqa: F401
    _XLSX_ENGINE: str = 'calamine'
except ImportError:
    _XLSX_ENGINE = 'openpyxl'

# ==============================================================================
# PRE-COMPILED REGEX  (compile once at import, reuse for every row)
# Note: these are all case-insensitive, and designed to pull relevant info from messy strings in various formats. They are not intended to be strict validators, but rather flexible extractors that can handle a variety of input formats commonly found in vulnerability and patch data.
# The _KB_RE is designed to extract KB article numbers from text, which are often used in Microsoft patch information. The _CVE_RE is designed to extract CVE identifiers, which follow a specific format of CVE-YYYY-NNNNN. The _VERSION_RE is designed to extract version numbers that may be in the format of major.minor.patch (e.g. 1.2.3) and can have between 2 and 5 components. The _DIGITS_RE is a simple regex to extract sequences of digits, which can be useful for parsing version numbers or other numeric information from text. The _NORM_CHARS regex is used to normalize text by replacing any sequence of non-alphanumeric characters with a single space, which can help with consistent text processing and comparison.
# By pre-compiling these regex patterns at the module level, we can improve the performance of our data processing functions, as they can reuse the compiled patterns instead of recompiling them for each row of data. This is especially beneficial when processing large datasets, as it can significantly reduce the overhead associated with regex operations.
# Overall, these pre-compiled regex patterns are essential tools for extracting and normalizing relevant information from the raw data we will be processing in our CVE dashboard, and they will help us to ensure that we can handle a wide variety of input formats effectively.  
# ==============================================================================

_KB_RE       = re.compile(r'KB\d+',                    re.IGNORECASE)
_CVE_RE      = re.compile(r'CVE-\d{4}-\d{4,7}',       re.IGNORECASE)
_VERSION_RE  = re.compile(r'\b\d+(?:\.\d+){1,4}\b')
_DIGITS_RE   = re.compile(r'\d+')
_NORM_CHARS  = re.compile(r'[^a-z0-9]+')

# get_base_product patterns
_ARCH_X64    = re.compile(r'\bx64\b',    re.IGNORECASE)
_ARCH_X86    = re.compile(r'\bx86\b',    re.IGNORECASE)
_ARCH_32     = re.compile(r'\b32-bit\b', re.IGNORECASE)
_ARCH_64     = re.compile(r'\b64-bit\b', re.IGNORECASE)
_EMPTY_PAREN = re.compile(r'\s*\(\s*\)')
_TRAILING_VER= re.compile(r'\s+v?\d[\d.+]*\s*$')
_SHEET_CHARS = re.compile(r'[\[\]\:\*\?\/\\\'\000]')

_CAT_COLS_VULN = ['Vulnerability Severity', 'Threat Status', 'Has Known Exploit', 'CISA KEV']
_CAT_COLS_RMM  = ['Device Type']


def _downcast_low_cardinality(df, cols):
    """Cast low-cardinality string columns to category dtype in-place."""
    for col in cols:
        if col in df.columns and not hasattr(df[col], 'cat'):
            df[col] = df[col].astype('category')
    return df


# RMM inventory column config (updateable via config.json)
_RMM_CFG = _CONFIG.get('rmm_inventory_columns', {})
_RMM_POSITIONAL = _RMM_CFG.get('positional_headers',
    ['Type','Client','Site','Device','Description','OS','Username','Last Response','Last Boot'])
_RMM_DEVICE_COL = _RMM_CFG.get('device_col', 'Device')
_RMM_RESP_COL   = _RMM_CFG.get('last_response_col', 'Last Response')
_RMM_OS_COL     = _RMM_CFG.get('os_col', 'OS')


# ==============================================================================
# PATCH GAP CLASSIFICATION
# This section includes functions and mappings to classify the gap between detected vulnerabilities and patch information. The classify_patch_gap function takes the patch match result and resolution status for a given vulnerability and determines if there is a gap in coverage, an unmanaged application, or a detection mismatch. The _GAP_NO_MATCH mapping is used to categorize specific no-match results from the patch matching process, while the _MATCHED_INSTALLED set is used to identify cases where the patch tool indicates that a patch is installed but the vulnerability is still detected by N-able, which may indicate a detection mismatch. By classifying these gaps, we can provide more actionable insights into the vulnerability management process and help users understand where there may be issues with patch coverage or detection accuracy.
# The classify_patch_gap function returns a specific gap category based on the patch match result and resolution status, which can then be used in the dashboard to highlight areas where there may be gaps in patch coverage or discrepancies between patch information and vulnerability detection. This helps users to prioritize their remediation efforts and focus on the most critical issues that may not be adequately addressed by their current patching strategy.
# The gap categories include 'coverage_gap' for vulnerabilities that are not found in the patch report, 'unmanaged_app' for devices that are in the patch report but the product is not found, and 'detection_mismatch' for cases where the patch tool indicates that a patch is installed but the vulnerability is still detected as unresolved. By identifying these gaps, users can take targeted actions to improve their vulnerability management and ensure that critical vulnerabilities are being effectively addressed.    
# Note: The classify_patch_gap function is designed to be flexible and can be easily extended in the future to include additional gap categories or handle new patch match results as needed. It provides a clear and structured way to classify the relationship between detected vulnerabilities and patch information, which is essential for effective vulnerability management and remediation prioritization. 
# ==============================================================================

# Maps Patch Match Result strings → explicit gap category (no-match cases).
_GAP_NO_MATCH: dict[str, str] = {
    'Not found in patch report':                  'coverage_gap',
    'Device in patch report - product not found': 'unmanaged_app',
}

# Matched-but-unresolved = patch tool says installed, N-able still detects CVE.
# These are the _STATUS_LABEL values for installed/reboot-required states.
_MATCHED_INSTALLED = {'Matched - installed', 'Matched - reboot required'}


def classify_patch_gap(patch_match_result: str,
                       resolved: Optional[str] = None) -> Optional[str]:
    """
    Return the explicit gap category for a patch row, or None if no gap.
    """
    pmr = str(patch_match_result).strip()

    if pmr in _GAP_NO_MATCH:
        return _GAP_NO_MATCH[pmr]

    if pmr in _MATCHED_INSTALLED and str(resolved).strip() == 'Unresolved':
        return 'detection_mismatch'

    return None
# ==============================================================================
# DATA LOADING & NORMALIZATION
# This section includes functions for loading data from CSV or Excel files, normalizing device names, extracting base product names, cleaning sheet names for Excel, extracting CVE IDs from text, determining device types based on OS strings, and parsing last response values into timestamps. These functions are essential for preparing the raw data for analysis and ensuring that it is in a consistent format for merging and processing in the dashboard generation. By normalizing device names and product names, we can improve the accuracy of merges between vulnerability data and RMM inventory data. The extraction functions help to pull relevant information from messy input strings, which is common in vulnerability and patch data. Overall, this section provides the necessary tools to transform raw input data into a structured format that can be effectively used in the subsequent analysis and dashboard generation steps.
# The load_data function can handle both CSV and Excel files, making it flexible for different input formats. The normalize_device_name function standardizes device names by stripping whitespace, converting to uppercase, and removing domain or subdomain components. The get_base_product function extracts a simplified product name by removing architecture tags and version information. The clean_sheet_name function ensures that sheet names for Excel are valid and unique. The extract_cve_id function pulls CVE identifiers from either raw strings or HYPERLINK formulas. The determine_device_type function classifies devices as 'Server', 'Workstation', or 'Unknown' based on their OS strings. The parse_last_response function converts various formats of last response values into sortable timestamps, which is crucial for analyzing the recency of vulnerability detections and responses.
# By using these functions to preprocess the data, we can ensure that the subsequent merging and analysis steps in the dashboard generation will be more accurate and effective, leading to a more insightful and actionable dashboard for users to understand their vulnerability landscape and prioritize remediation efforts.
# Note: These functions are designed to be robust and handle a variety of input formats and edge cases, which is important given the often messy nature of vulnerability and inventory data. By implementing these normalization and extraction functions, we can improve the quality of our data and the insights derived from it in the CVE dashboard.
# ==============================================================================
def load_data(file_path: str) -> pd.DataFrame:
    if file_path.lower().endswith(('.xlsx', '.xls')):
        return pd.read_excel(file_path, engine=_XLSX_ENGINE)
    return pd.read_csv(file_path)
# The load_data function is designed to handle both CSV and Excel file formats, making it flexible for different types of input data. It checks the file extension to determine whether to use pandas' read_excel or read_csv function to load the data into a DataFrame. This allows users to provide their vulnerability and inventory data in the format that is most convenient for them, without needing to convert it beforehand. By returning a DataFrame, this function provides a consistent data structure for subsequent processing steps in the dashboard generation.
# Note: The function assumes that the input file is well-formed and does not include error handling for cases such as missing files, unsupported formats, or malformed data. In a production environment, it would be advisable to add error handling to provide more informative feedback to the user in case of issues with the input file.
def normalize_device_name(name: str) -> str:
    """Row-level device name normalisation (used for single values)."""
    name = str(name).strip().upper()
    if '\\' in name: name = name.split('\\')[-1]
    if '.'  in name: name = name.split('.')[0]
    return name

# The _normalize_device_col function is a vectorized version of the normalize_device_name function, designed to operate on entire pandas Series (columns) of device names. It applies the same normalization logic (stripping whitespace, converting to uppercase, and removing domain or subdomain components) to each value in the Series efficiently using pandas string methods. This allows for fast processing of large datasets without the need for explicit loops, which can be slow in Python. By normalizing the device names in a vectorized manner, we can prepare the data for merging and analysis more quickly, improving the overall performance of the dashboard generation process.
# Note: This function assumes that the input Series contains string-like values and will convert any non-string values to strings before applying the normalization. It also handles cases where the expected delimiters ('\\' and '.') may not be present, ensuring that it can process a wide variety of input formats without raising errors. By using this function to normalize device names in the RMM inventory data, we can improve the accuracy of merges with vulnerability data, which is crucial for correctly associating vulnerabilities with the devices they affect in the dashboard.
# The normalization process helps to ensure that device names are consistent across different data sources, which is essential for accurate merging and analysis. By removing domain components and standardizing the case, we can reduce the likelihood of mismatches due to formatting differences, leading to a more accurate and insightful dashboard for users to understand their vulnerability landscape and prioritize remediation efforts effectively.
# Overall, the _normalize_device_col function is a key component of the data preprocessing pipeline, enabling efficient and consistent normalization of device names across large datasets, which is critical for the success of the CVE dashboard generation.
def _normalize_device_col(series: 'pd.Series') -> 'pd.Series':
    """
    Vectorised version of normalize_device_name for DataFrame columns.
    """
    s = series.astype(str).str.strip().str.upper()
    s = s.str.split('\\\\').str[-1]
    s = s.str.split('\\.').str[0]
    return s
# The get_base_product function is designed to extract a simplified base product name from a given product name string. It removes common architecture tags (such as x64, x86, 32-bit, 64-bit), empty parentheses, and trailing version information to distill the core product name. This is useful for grouping and analyzing vulnerabilities by product without being confounded by variations in how the product name may be formatted in different data sources. By extracting the base product name, we can improve the accuracy of merges between vulnerability data and patch information, as well as enhance the clarity of the dashboard by presenting more concise product names.
# The function uses pre-compiled regular expressions to efficiently remove the unwanted components from the product name. It also ensures that the resulting base product name is stripped of leading and trailing whitespace. This helps to standardize product names across different data sources, which is crucial for accurate analysis and reporting in the CVE dashboard. By using this function to extract base product names, we can provide clearer insights into which products are affected by vulnerabilities and how they relate to available patches and remediation efforts.
# Note: The get_base_product function is designed to be flexible and can handle a variety of input formats for product names. However, it may not cover all possible variations, and there may be cases where additional rules or exceptions are needed to accurately extract the base product name. In such cases, the function can be easily extended with additional regex patterns or logic to handle specific edge cases as they arise in the data.
def get_base_product(prod_name: str) -> str:
    p = str(prod_name).strip()
    p = _ARCH_X64.sub('', p)
    p = _ARCH_X86.sub('', p)
    p = _ARCH_32.sub('', p)
    p = _ARCH_64.sub('', p)
    p = _EMPTY_PAREN.sub('', p)
    p = _TRAILING_VER.sub('', p)
    return p.strip()
# The clean_sheet_name function is designed to take a raw product name and convert it into a valid Excel sheet name. It removes or replaces characters that are not allowed in sheet names, truncates the name to a maximum of 31 characters, and ensures that the resulting sheet name is unique within the context of already used names. If the input name is empty or consists only of whitespace, it defaults to 'Unknown Product'. The function also handles cases where multiple products may have similar names by appending a numeric suffix to create unique sheet names. This is important for ensuring that the generated Excel file can be created without errors due to invalid or duplicate sheet names, and it helps to maintain clarity and organization in the resulting workbook.
# By using the clean_sheet_name function to generate sheet names for each product in the dashboard,
# we can ensure that the Excel output is well-structured and free of naming issues, which enhances the usability of the dashboard for users who may want to explore the data in Excel. The function's ability to handle edge cases and ensure uniqueness of sheet names contributes to a more robust and user-friendly output.
# Note: The clean_sheet_name function assumes that the input name is a string and will convert non-string inputs to strings. It also uses a set of already used names to ensure uniqueness, which should be maintained and passed correctly when generating sheet names for multiple products. By implementing this function, we can avoid common pitfalls with Excel sheet naming and provide a smoother experience for users interacting with the generated dashboard in Excel.
def clean_sheet_name(name: str, used_names: Set[str]) -> str:
    if pd.isna(name) or str(name).strip() == '': name = 'Unknown Product'
    clean = _SHEET_CHARS.sub('', str(name)).strip()[:31].strip()
    if not clean: clean = 'Unknown Product'
    final, counter = clean, 1
    while final.lower() in {n.lower() for n in used_names}:
        suffix = f'_{counter}'
        final = clean[:31 - len(suffix)] + suffix
        counter += 1
    used_names.add(final)
    return final
# The extract_cve_id function is designed to pull a bare CVE identifier (in the format CVE-YYYY-NNNNN) from a given string, which may be a raw string or a HYPERLINK formula commonly found in Excel. It uses the pre-compiled CVE_PATTERN regex to search for the CVE ID within the input string. If a match is found, it returns the CVE ID in uppercase; otherwise, it returns the original input string stripped of whitespace and converted to uppercase. This function is essential for standardizing CVE identifiers across different data sources, which may have varying formats for how CVEs are represented. By extracting and normalizing CVE IDs, we can improve the accuracy of merges and analyses that rely on CVE identifiers in the dashboard.
# The function is designed to be flexible and can handle cases where the input string may contain additional text or formatting around the CVE ID, as it specifically looks for the CVE pattern within the string. This helps to ensure that we can accurately extract CVE IDs even from messy input data, which is common in vulnerability reports and spreadsheets. By using this function to extract CVE IDs, we can ensure that our data is consistent and that we can effectively link vulnerabilities to their corresponding CVE identifiers in the dashboard, providing clearer insights into the vulnerabilities being analyzed and their associated information.
#  Note: The extract_cve_id function assumes that the input string may contain a CVE ID in the expected format, and it will return the first match found. If there are multiple CVE IDs in the input string, it will only return the first one. In cases where the input string does not contain a valid CVE ID, it will return the original string in uppercase, which may not be ideal for all use cases. Depending on the context in which this function is used, additional error handling or validation may be necessary to ensure that the output is appropriate for the intended use in the dashboard.        
def extract_cve_id(val: str) -> str:
    """Pull a bare CVE-YYYY-NNNNN from either a raw string or a HYPERLINK formula."""
    m = CVE_PATTERN.search(str(val))
    return m.group(1).upper() if m else str(val).strip().upper()
# The determine_device_type function classifies a device as 'Server', 'Workstation', or 'Unknown' based on the content of its OS string. It checks for specific keywords in the OS string to make this determination. If the OS string contains 'server', it classifies the device as a 'Server'. If it contains 'windows 10' or 'windows 11', it classifies it as a 'Workstation'. If the OS string is 'nan' or 'unknown', it classifies it as 'Unknown'. For any other cases, it defaults to classifying the device as a 'Workstation'. This function helps to categorize devices in the RMM inventory, which can be useful for analyzing vulnerabilities and patch information by device type in the dashboard.
# The function is designed to be simple and relies on keyword matching, which may not cover all possible OS strings or device types. However, it provides a basic classification that can be useful for many common cases. Depending on the specific OS strings encountered in the data, additional rules or keywords may need to be added to improve the accuracy of the classification. By using this function to determine device types, we can enhance the insights provided in the dashboard by allowing users to filter and analyze vulnerabilities and patches based on whether they affect servers or workstations.
# Note: The determine_device_type function assumes that the input OS string is a string-like value and will convert non-string inputs to strings. It also uses simple keyword matching, which may not be sufficient for all cases, especially if there are variations in how OS information is represented in the data. In a production environment, it may be beneficial to implement a more robust classification system that can handle a wider variety of OS strings and device types, potentially using machine learning or more complex rule-based logic if necessary. However, for many common cases, this function provides a straightforward way to classify devices based on their OS information.
def determine_device_type(os_string: str) -> str:
    val = str(os_string).lower()
    if val in ('nan', 'unknown'): return 'Unknown'
    if 'server' in val: return 'Server'
    if 'windows 10' in val or 'windows 11' in val: return 'Workstation'
    return 'Workstation'

# The parse_last_response function converts Last Response values into sortable timestamps. It handles various formats of input values, including special cases like 'Not Found in RMM', 'N/A', and empty strings. For valid date strings, it attempts to convert them to pandas datetime objects. If the conversion fails, it tries to parse the value based on specific patterns, such as 'overdue_' prefixes or time-based descriptions containing 'days' or 'hrs'. If all else fails, it returns a default epoch timestamp. This function is crucial for analyzing the recency of vulnerability detections and responses, allowing us to sort and filter data based on when the last response was recorded. By converting various formats of last response values into a consistent timestamp format, we can enhance the analysis capabilities in the dashboard and provide more actionable insights into the timeliness of responses to detected vulnerabilities.
# The function is designed to be robust and handle a variety of input formats, which is important given the often inconsistent nature of data in RMM inventory reports. By providing a consistent way to parse last response values, we can ensure that our analyses of response times and trends are based on accurate and comparable data, which is essential for effective vulnerability management and remediation prioritization in the dashboard. Note that the function assumes that the input values may be in various formats and includes logic to handle common cases, but there may still be edge cases that require additional handling as they are encountered in real-world data.
def parse_last_response(val):
    """Parse Last Response values into sortable timestamps."""
    val = str(val).strip()
    epoch = pd.to_datetime('1900-01-01')
    if val in ['Not Found in RMM', 'N/A', '']: return epoch
    try: return pd.to_datetime(val)
    except Exception: pass
    if val.startswith('overdue_'):
        try: return pd.to_datetime(val.replace('overdue_', '').split(' -')[0])
        except Exception: pass
    if 'days' in val or 'hrs' in val:
        try:
            m = _DIGITS_RE.search(val)
            days = int(m.group(0)) if m else 0
            return pd.Timestamp.now() - pd.Timedelta(days=days)
        except Exception: pass
    return epoch
# The get_col_letter function converts a zero-based column index into an Excel column letter. It uses a loop to repeatedly divide the column index by 26 and determine the corresponding letter for each position. This is useful for generating Excel formulas or references that require column letters instead of numeric indices. By providing this utility function, we can easily convert between numeric column indices used in pandas DataFrames and the letter-based column references used in Excel, which is essential for creating accurate formulas and references in the generated dashboard when exporting to Excel. The function handles the conversion correctly by accounting for the fact that Excel columns are 1-indexed and that after 'Z' comes 'AA', 'AB', etc. By using this function, we can ensure that any Excel-related operations in the dashboard generation can reference columns accurately, regardless of how many columns are present in the data.
def get_col_letter(col_idx):
    letter = ''
    col_idx += 1
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        letter = chr(65 + remainder) + letter
    return letter
# The _drop_internal function is designed to remove columns that are used internally within the data processing pipeline but should not be included in the final output when writing to Excel. It takes a DataFrame as input and drops any columns that are listed in the specified set of internal column names, if they exist in the DataFrame. This helps to ensure that the final Excel output is clean and only includes relevant information for the end user, without exposing any intermediate columns that were used for merging, sorting, or other internal operations during the data processing. By using this function before writing to Excel, we can maintain a clear separation between the internal workings of the data pipeline and the final output presented to users in the dashboard.
# The function uses the pandas drop method with errors='ignore' to avoid raising an error if any of the specified internal columns are not present in the DataFrame. This allows for flexibility in the data processing pipeline, as different steps may add or remove internal columns as needed. By centralizing the
def _drop_internal(df):
    """Drop pipeline-only columns before writing to Excel."""
    return df.drop(columns=[c for c in ('Name_Join', 'Device_Join', 'Base Product',
                                         '_Sort_Time', '_Name_Key', '_CVE_Key',
                                         '_Checkbox_Resolved')
                             if c in df.columns], errors='ignore').copy()


# ==============================================================================
# PATCH MATCH HELPER FUNCTIONS
# This section includes helper functions for patch matching and classification, such as normalizing text for comparison, extracting architecture information from product names, mapping vulnerability names to products, extracting KB articles and CVE IDs from text, determining the best version from a string, parsing version numbers for comparison, and classifying baseline compliance and version check results. These functions are essential for accurately matching detected vulnerabilities with available patches and determining the compliance status of devices based on their installed patches and the known fixed versions for vulnerabilities. By implementing these helper functions, we can enhance the accuracy and effectiveness of the patch matching process in the dashboard generation, providing users with clearer insights into their vulnerability landscape and remediation efforts. The functions are designed to handle a variety of input formats and edge cases, which is important given the
# often messy nature of vulnerability and patch data. By using these functions in the patch matching process, we can ensure that we are accurately identifying which patches correspond to which vulnerabilities and providing meaningful classifications of compliance and resolution status for users to act upon in the dashboard.
# Note: The functions in this section are intended to be used as part of the patch matching and classification process, and they may rely on specific formats or conventions in the input data. It is important to ensure that the data being processed is compatible with the expectations of these functions, and additional error handling or validation may be necessary in a production environment to handle unexpected input formats or edge cases. By carefully implementing and using these helper functions, we can improve the overall quality and usefulness of the CVE dashboard for users to understand their vulnerability landscape and prioritize remediation efforts effectively.
# ==============================================================================
# The _norm_compact function normalizes a string by removing all non-alphanumeric characters, converting it to lowercase, and stripping leading and trailing whitespace. This is useful for creating a compact version of a string that can be used for comparison or matching purposes, as it reduces the variability in formatting and allows for more consistent comparisons between strings that may have different punctuation or spacing. By using this function to normalize product names, vulnerability names, or other relevant text fields, we can improve the accuracy of matching and classification in the patch matching process. The _norm_text function is similar but replaces sequences of non-alphanumeric characters with a single space instead of removing them entirely, which can be useful for preserving some separation between words while still normalizing the text for comparison. Both functions help to ensure that we can effectively compare and match strings in a way that is robust to common formatting differences and inconsistencies in the input data, which is essential for accurate analysis and reporting in the CVE dashboard.
def _norm_compact(v): return _NORM_CHARS.sub('', str(v).lower()).strip()
def _norm_text(v):    return _NORM_CHARS.sub(' ', str(v).lower()).strip()

_ARCH_TAG_RE = re.compile(r'[(](x64|x86|32[\-\s]?bit|64[\-\s]?bit)[)]', re.IGNORECASE)
# The _get_arch function uses a regular expression to search for architecture tags in a given text string, such as '(x64)', '(x86)', '(32-bit)', or '(64-bit)'. If a match is found, it returns 'x86' for 32-bit architectures and 'x64' for 64-bit architectures. If no match is found, it returns an empty string. This function is useful for extracting architecture information from product names or descriptions, which can be important for accurately matching vulnerabilities to patches that may be specific to certain architectures. By using this function in the patch matching process, we can ensure that we are correctly identifying the relevant patches for each vulnerability based on the architecture of the affected product, which enhances the accuracy and usefulness of the dashboard for users to understand their vulnerability landscape and prioritize remediation efforts effectively.
# Note: The _get_arch function assumes that the architecture information is enclosed in parentheses and follows the specific formats mentioned. If there are variations in how architecture information is represented in the input data, additional patterns may need to be added to the regular expression to ensure accurate extraction. Additionally, the function currently only distinguishes between x86 and x64 architectures, so if there are other architectures that need to be identified, further logic may be necessary to handle those cases.
def _get_arch(text: str) -> str:
    m = _ARCH_TAG_RE.search(str(text))
    if not m:
        return ''
    a = m.group(1).lower()
    return 'x86' if ('x86' in a or '32' in a) else 'x64'

STATUS_RANK = {'Installed': 6, 'Reboot Required': 5, 'Installing': 4,
                'Pending': 3, 'Missing': 2, 'Failed': 1}
STATUS_LABEL = {
    'Installed':       'Matched - installed',
    'Reboot Required': 'Matched - reboot required',
    'Installing':      'Matched - installing',
    'Pending':         'Matched - pending',
    'Missing':         'Matched - missing',
    'Failed':          'Matched - failed',
}
INSTALLED_STATUSES = {'Installed', 'Reboot Required'}

def _detect_product(text):
    t = _norm_text(str(text))
    for key, product in PRODUCT_MAP:
        if _norm_text(key) in t: return product
    return ''

def _extract_kbs(text) -> list:
    return sorted({kb.upper() for kb in _KB_RE.findall(str(text))})

def _extract_cves(text) -> list:
    return sorted({c.upper() for c in _CVE_RE.findall(str(text))})

def _extract_best_version(text) -> str:
    versions = _VERSION_RE.findall(str(text))
    if not versions: return ''
    return sorted(versions, key=lambda v: (len(v.split('.')), [int(x) for x in v.split('.')]))[-1]

def _parse_version(value) -> Optional[tuple]:
    parts = _DIGITS_RE.findall(str(value).strip())
    return tuple(int(p) for p in parts) if parts else None

def _version_gte(left, right):
    l, r = _parse_version(left), _parse_version(right)
    if l is None or r is None: return None
    n = max(len(l), len(r))
    return (l + (0,) * (n - len(l))) >= (r + (0,) * (n - len(r)))

def _make_excel_safe(df):
    out = df.copy()
    for col in out.columns:
        if isinstance(out[col].dtype, pd.DatetimeTZDtype):
            out[col] = out[col].dt.tz_localize(None)
    return out

def _resolve_fixed_version(row):
    if 'Fixed Version' in row.index:
        v = str(row.get('Fixed Version', '')).strip()
        if v: return v, 'CVE workbook column'
    product = row.get('_pk', '')
    if not product: return '', ''
    rules = FIXED_VERSION_RULES.get(product, {})
    for cve in _extract_cves(str(row.get('Vulnerability Name', ''))):
        if cve in rules: return rules[cve], f'config rule ({cve})'
    return '', ''

def _resolve_baseline(row) -> tuple[str, str]:
    product = row.get('_pk', '')
    if not product: return '', ''
    rules = FIXED_VERSION_RULES.get(product, {})
    baseline = rules.get('_baseline', '').strip()
    if baseline: return baseline, 'rolling baseline'
    return '', ''


def _classify_baseline_compliance(row) -> str:
    status = str(row.get('Status', '')).strip()
    if status not in INSTALLED_STATUSES:
        return 'Not installed'
    bl = str(row.get('Product Baseline', '')).strip()
    if not bl:
        return 'No baseline defined'
    pv = str(row.get('Matched Patch Version', '')).strip()
    if not pv:
        return 'Version unknown'
    cmp = _version_gte(pv, bl)
    if cmp is True:  return 'Compliant'
    if cmp is False: return 'Below baseline'
    return 'Version unknown'


def _classify_version_check(row):
    status = str(row.get('Status', '')).strip()
    pv     = str(row.get('Matched Patch Version', '')).strip()
    fv     = str(row.get('Fixed Version Used', '')).strip()
    if status not in INSTALLED_STATUSES:
        return 'Patch not yet installed' if status else 'No patch evidence'
    if not fv:
        return 'Installed version found - no fixed baseline' if pv else 'Installed - version unknown'
    if not pv: return 'Fixed baseline known - installed version not found'
    cmp = _version_gte(pv, fv)
    if cmp is True:  return 'Version compliant'
    if cmp is False: return 'Below fixed version'
    return 'Version comparison failed'

def _classify_resolution(row):
    status = str(row.get('Status', '')).strip()
    if status not in INSTALLED_STATUSES:
        return 'Unresolved'

    vcr = str(row.get('Version Check Result', '')).strip().lower()

    if 'below fixed version' in vcr:
        return 'Unresolved'

    if 'no fixed baseline' in vcr:
        return 'Unresolved'

    if 'version compliant' not in vcr:
        return 'Unresolved'

    try:
        install_dt = pd.to_datetime(row.get('Patch Install Date'), errors='coerce')
        if pd.isna(install_dt):
            return 'Unresolved'

        cve_dates = []
        for col in ('First detected', 'Date Published'):
            v = row.get(col)
            if v is not None and not (isinstance(v, float) and pd.isna(v)):
                dt = pd.to_datetime(v, errors='coerce')
                if not pd.isna(dt):
                    cve_dates.append(dt)

        if not cve_dates:
            return 'Unresolved'

        return 'Patch confirmed - pending rescan' if install_dt >= max(cve_dates) else 'Unresolved'

    except Exception:
        return 'Unresolved'


# ==============================================================================
# DATA PIPELINE: CVE + RMM
# ==============================================================================

def load_vulnerability_data(file_path: str) -> pd.DataFrame:
    if str(file_path).lower().endswith(('.xlsx', '.xls')):
        xl = pd.ExcelFile(file_path, engine=_XLSX_ENGINE)
        if 'Raw Data' in xl.sheet_names:
            log.info("Detected dashboard workbook — reading 'Raw Data' sheet")
            df = xl.parse('Raw Data')
        else:
            df = xl.parse(xl.sheet_names[0])
    else:
        try:
            df = pd.read_csv(file_path, dtype_backend='pyarrow')
        except TypeError:
            df = pd.read_csv(file_path)

    rename = {}
    for col in df.columns:
        c = str(col).strip().lower()
        if   c in ('asset name', 'device name', 'endpoint'):
            rename[col] = 'Name'
        elif c in ('vulnerability id', 'cve id', 'cve'):
            rename[col] = 'Vulnerability Name'
        elif c in ('cvss score', 'cvss v3.1 base score', 'cvss v3 base score',
                   'base score', 'score'):
            rename[col] = 'Vulnerability Score'
        elif c in ('affected products', 'product'):
            rename[col] = 'Affected Products'
        elif c in ('severity', 'risk'):
            rename[col] = 'Vulnerability Severity'
        elif c in ('threat status',):
            rename[col] = 'Threat Status'
        elif c in ('customer name', 'client name', 'account name', 'client'):
            rename[col] = 'Customer'
        elif c in ('site name', 'location name'):
            rename[col] = 'Site'
        elif c in ('has exploit', 'exploit') and 'Has Known Exploit' not in df.columns:
            rename[col] = 'Has Known Exploit'
        elif c == 'cisa kev' and col != 'CISA KEV' and 'CISA KEV' not in df.columns:
            rename[col] = 'CISA KEV'
        elif c == 'first detected' and col != 'First detected':
            rename[col] = 'First detected'
        elif c == 'last updated' and col != 'Last updated':
            rename[col] = 'Last updated'
        elif c in ('updates available', 'update available'):
            rename[col] = 'Update Available'
        elif c in ('operating system role', 'os role'):
            rename[col] = 'Operating System Role'
    df.rename(columns=rename, inplace=True)

    defaults = {
        'Name': 'Unknown Device',          'Vulnerability Name': 'Unknown CVE',
        'Affected Products': 'Unknown Product', 'Vulnerability Score': 0.0,
        'Vulnerability Severity': 'Unknown',    'Has Known Exploit': 'No',
        'CISA KEV': 'No',                       'Risk Severity Index': 'Unknown',
    }
    for col, default in defaults.items():
        if col not in df.columns: df[col] = default

    df['Vulnerability Name'] = df['Vulnerability Name'].fillna('Unknown CVE')
    df['Name_Join']          = _normalize_device_col(df['Name'])
    df['Affected Products']  = df['Affected Products'].fillna('Unknown Product')
    df['Base Product']       = df['Affected Products'].apply(get_base_product)

    _downcast_low_cardinality(df, _CAT_COLS_VULN)

    return df

def load_rmm_data(file_path):
    df        = load_data(file_path)
    col_lower = {c.lower(): c for c in df.columns}
    dev_col = resp_col = os_col = device_type_col = None

    for key in ('device name', 'device', 'name', 'asset name', 'hostname'):
        if key in col_lower: dev_col = col_lower[key]; break
    for key in ('last response (local time)', 'last response (utc)', 'last response', 'last check-in'):
        if key in col_lower: resp_col = col_lower[key]; break
    for key in ('os version', 'os'):
        if key in col_lower: os_col = col_lower[key]; break
    for key in ('device type',):
        if key in col_lower: device_type_col = col_lower[key]; break

    username_col = None
    for key in ('username', 'user name', 'logged-on user', 'logged on user',
                'current user', 'user'):
        if key in col_lower:
            username_col = col_lower[key]
            break

    if not dev_col or not resp_col:
        cols_are_positional = all(
            isinstance(c, int) or str(c).startswith('Unnamed')
            for c in df.columns
        )
        if len(df.columns) == len(_RMM_POSITIONAL) and cols_are_positional:
            df.columns = _RMM_POSITIONAL
            dev_col, resp_col, os_col = _RMM_DEVICE_COL, _RMM_RESP_COL, _RMM_OS_COL
        else:
            raise ValueError(
                "Could not identify required columns in RMM/Device Inventory file.\n\n"
                f"Looking for:  '{_RMM_DEVICE_COL}' and '{_RMM_RESP_COL}'.\n"
                f"Found columns: {', '.join(str(c) for c in df.columns[:12])}"
                + (' ...' if len(df.columns) > 12 else '') + "\n\n"
                "To fix without code changes, update 'rmm_inventory_columns' in config.json."
            )

    rename_map = {dev_col: 'Device', resp_col: 'Last Response'}
    if username_col and username_col != 'Username':
        rename_map[username_col] = 'Username'
    df.rename(columns=rename_map, inplace=True)
    if 'Username' not in df.columns:
        df['Username'] = ''
    df['Username'] = df['Username'].fillna('').astype(str).str.strip()
    df['Device_Join'] = _normalize_device_col(df['Device'])

    if device_type_col:
        def _map_device_type(val):
            v = str(val).strip().upper()
            if v == 'SERVER':  return 'Server'
            if v in ('', 'NAN', 'UNKNOWN'): return 'Unknown'
            return 'Workstation'
        df['Device Type'] = df[device_type_col].apply(_map_device_type)
    elif os_col:
        df['Device Type'] = df[os_col].apply(determine_device_type)
    else:
        df['Device Type'] = 'Unknown'

    _downcast_low_cardinality(df, _CAT_COLS_RMM)

    return df.drop_duplicates(subset=['Device_Join'], keep='first')

def merge_data(df_vuln, df_rmm, skip_rmm, exclude_missing_rmm=True):
    vuln_has_lr = 'Last Response' in df_vuln.columns
    vuln_has_dt = 'Device Type'   in df_vuln.columns

    if not skip_rmm and df_rmm is not None:
        rmm_pull = ['Device_Join']
        if not vuln_has_lr: rmm_pull.append('Last Response')
        if not vuln_has_dt: rmm_pull.append('Device Type')
        # Always pull Username when it exists in the RMM frame
        if 'Username' in df_rmm.columns and 'Username' not in df_vuln.columns:
            rmm_pull.append('Username')

        if len(rmm_pull) > 1:
            before = len(df_vuln)
            join_how = 'inner' if exclude_missing_rmm else 'left'
            merged = pd.merge(df_vuln, df_rmm[rmm_pull],
                              left_on='Name_Join', right_on='Device_Join', how=join_how)

            # Widen category columns to object before any .loc writes.
            for _cat_col in ('Device Type', 'Last Response'):
                if _cat_col in merged.columns and hasattr(merged[_cat_col], 'cat'):
                    merged[_cat_col] = merged[_cat_col].astype(object)

            dropped = before - len(merged)
            if dropped and exclude_missing_rmm:
                decom_names = (
                    set(df_vuln['Name_Join'].unique())
                    - set(df_rmm['Device_Join'].unique())
                )
                log.info(
                    "Excluded %d CVE rows for %d decommissioned device(s) "
                    "(not in Device Inventory): %s%s",
                    dropped, len(decom_names),
                    ', '.join(sorted(decom_names)[:5]),
                    ' ...' if len(decom_names) > 5 else '',
                )
            elif not exclude_missing_rmm:
                missing_mask = merged['Device_Join'].isna()
                if missing_mask.any():
                    if not vuln_has_lr:
                        merged.loc[missing_mask, 'Last Response'] = 'Not Found in RMM'
                    if not vuln_has_dt:
                        merged.loc[missing_mask, 'Device Type'] = 'Unknown'
                    log.info(
                        "%d CVE rows for devices not in Device Inventory kept "
                        "(exclude_missing_rmm=False)",
                        missing_mask.sum(),
                    )
            if not vuln_has_dt:
                merged['Device Type'] = merged['Device Type'].fillna('Unknown')
        else:
            merged = df_vuln.copy()
    else:
        merged = df_vuln.copy()
        if not vuln_has_lr:
            merged['Last Response'] = 'N/A'
        if not vuln_has_dt:
            if 'Operating System Role' in merged.columns:
                merged['Device Type'] = merged['Operating System Role'].str.title()
            else:
                merged['Device Type'] = 'Unknown'

    if 'Last Response' not in merged.columns: merged['Last Response'] = 'N/A'
    if 'Device Type'   not in merged.columns: merged['Device Type']   = 'Unknown'
    if 'Username'      not in merged.columns: merged['Username']      = ''
    merged['Username'] = merged['Username'].fillna('').astype(str).str.strip()

    if 'Operating System Role' in merged.columns:
        _OS_ROLE_MAP = {'WORKSTATION': 'Workstation', 'SERVER': 'Server'}
        mask_unk = merged['Device Type'].astype(str).str.strip().str.lower() == 'unknown'
        merged.loc[mask_unk, 'Device Type'] = (
            merged.loc[mask_unk, 'Operating System Role']
            .astype(str).str.strip().str.upper()
            .map(_OS_ROLE_MAP)
            .fillna('Unknown')
        )

    if 'OS' in merged.columns:
        mask_still_unk = merged['Device Type'].astype(str).str.strip().str.lower() == 'unknown'
        if mask_still_unk.any():
            def _infer_from_os(val):
                v = str(val).lower()
                if 'server' in v:     return 'Server'
                if 'windows' in v:    return 'Workstation'
                return 'Unknown'
            merged.loc[mask_still_unk, 'Device Type'] = (
                merged.loc[mask_still_unk, 'OS'].apply(_infer_from_os)
            )

    merged['Vulnerability Score'] = pd.to_numeric(merged['Vulnerability Score'], errors='coerce')

    # ── Vectorised _Sort_Time + Days Since Last Response ─────────────────────
    # Replaces two row-by-row .apply(parse_last_response) loops with bulk
    # pd.to_datetime calls.  format='mixed' silences the inference warning and
    # handles the mix of date formats N-able exports produce.
    _epoch         = pd.to_datetime('1900-01-01')
    _lr_str        = merged['Last Response'].astype(str).str.strip()
    _sentinel_mask = _lr_str.isin(['Not Found in RMM', 'N/A', ''])

    _sort_time = pd.to_datetime(
        _lr_str.where(~_sentinel_mask, other=pd.NaT),
        errors='coerce',
        format='mixed',
        dayfirst=False,
    )
    if hasattr(_sort_time, 'dt') and _sort_time.dt.tz is not None:
        _sort_time = _sort_time.dt.tz_localize(None)
    _sort_time = _sort_time.fillna(_epoch)
    merged['_Sort_Time'] = _sort_time

    # Fallback: rows still at epoch → try publication/detection date columns
    _stale_mask = merged['_Sort_Time'] <= _epoch
    if _stale_mask.any():
        for _date_col in ('Last updated', 'First detected', 'Date Published'):
            if _date_col in merged.columns:
                _parsed = pd.to_datetime(
                    merged[_date_col].astype(str).str.replace(' UTC', '', regex=False),
                    errors='coerce',
                    utc=True,
                ).dt.tz_localize(None)
                _update_mask = _stale_mask & _parsed.notna()
                merged.loc[_update_mask, '_Sort_Time'] = _parsed[_update_mask]
                _stale_mask = merged['_Sort_Time'] <= _epoch
                if not _stale_mask.any():
                    break

    # Days Since Last Response — reuse already-parsed series, no second parse
    _now      = pd.Timestamp.now()
    _no_data  = _sentinel_mask | (merged['_Sort_Time'] <= _epoch)
    _days_num = (_now - merged['_Sort_Time']).dt.days.clip(lower=0).astype(object)
    _days_num[_no_data] = '—'
    merged['Days Since Last Response'] = _days_num
    
    cols = merged.columns.tolist()
    if 'Days Since Last Response' in cols:
        cols.remove('Days Since Last Response')
        lr_idx = cols.index('Last Response') if 'Last Response' in cols else len(cols)
        cols.insert(lr_idx + 1, 'Days Since Last Response')
        merged = merged[cols]

    _downcast_low_cardinality(merged, _CAT_COLS_VULN + _CAT_COLS_RMM)

    return merged


# ==============================================================================
# DATA PIPELINE: PATCH MATCH
# ==============================================================================

def _apply_cascade_resolution(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or 'Matched Patch Version' not in df.columns:
        return df

    df = df.copy()

    df['_cascade_pk'] = df.get('_pk', df['Affected Products'].apply(
        lambda v: _detect_product(str(v).lower())))

    installed_mask = df['Patch Match Result'].astype(str).str.strip().isin(
        {'Matched - installed', 'Matched - reboot required'}
    )
    has_ver = df[installed_mask & 
                 (df['Matched Patch Version'].astype(str).str.strip().str.len() > 2)].copy()
    has_ver['_vt'] = has_ver['Matched Patch Version'].apply(_parse_version)
    has_ver = has_ver[has_ver['_vt'].notna()]
    if has_ver.empty:
        return df

    best_ver: dict[tuple, dict] = {}
    for _, row in has_ver.iterrows():
        _patch_arch = _get_arch(str(row.get('Matched Patch', '')))
        key = (str(row['Name']), str(row['_cascade_pk']), _patch_arch)
        vt  = row['_vt']
        if key not in best_ver or vt > best_ver[key]['_vt']:
            best_ver[key] = {
                '_vt':          vt,
                'version_str':  str(row['Matched Patch Version']),
                'install_date': row.get('Patch Install Date', pd.NaT),
                '_arch':        _patch_arch,
            }

    cascade_resolve: set[tuple[str, str, str]] = set()

    product_baselines: dict[str, tuple] = {}
    for pk, rules in FIXED_VERSION_RULES.items():
        if isinstance(rules, dict) and '_baseline' in rules:
            bt = _parse_version(rules['_baseline'])
            if bt:
                product_baselines[pk] = bt

    for (device, pk, inst_arch), ver_info in best_ver.items():
        rules = FIXED_VERSION_RULES.get(pk, {})
        if not isinstance(rules, dict):
            continue

        for cve_id, fixed_str in rules.items():
            if cve_id.startswith('_'):
                continue
            fixed_t = _parse_version(fixed_str)
            if fixed_t and ver_info['_vt'] >= fixed_t:
                cascade_resolve.add((device, pk, cve_id.upper(), inst_arch))

        if pk in product_baselines and ver_info['_vt'] >= product_baselines[pk]:
            cascade_resolve.add((device, pk, '_BASELINE_', inst_arch))

    if not cascade_resolve:
        return df

    cascade_applied = 0
    for idx, row in df.iterrows():
        if str(row.get('Patch Evidence Status', '')).strip() != 'Unresolved':
            continue
        device  = str(row['Name'])
        pk      = str(row.get('_cascade_pk', ''))
        cve_ids = [c.upper() for c in _extract_cves(str(row.get('Vulnerability Name', '')))]

        cve_arch = _get_arch(str(row.get('Affected Products', '')))

        key_exact   = (device, pk, cve_arch)
        key_neutral = (device, pk, '')
        if key_exact in best_ver:
            key = key_exact
        elif key_neutral in best_ver and not cve_arch:
            key = key_neutral
        elif not cve_arch and any(k[:2] == (device, pk) for k in best_ver):
            key = next(k for k in best_ver if k[:2] == (device, pk))
        else:
            continue

        ver_info   = best_ver[key]
        install_dt = pd.to_datetime(ver_info['install_date'], errors='coerce')
        first_dt   = pd.to_datetime(row.get('First detected', pd.NaT), errors='coerce')
        inst_arch  = ver_info['_arch']

        for cve_id in cve_ids:
            if (device, pk, cve_id, inst_arch) in cascade_resolve or \
               (not cve_arch and any((device, pk, cve_id, a) in cascade_resolve for a in ['', 'x64', 'x86'])):
                if not pd.isna(install_dt) and not pd.isna(first_dt) and install_dt >= first_dt:
                    df.at[idx, 'Patch Evidence Status'] = 'Patch confirmed - pending rescan'
                    cascade_applied += 1
                    break
            elif (device, pk, '_BASELINE_', inst_arch) in cascade_resolve or \
                   (not cve_arch and any((device, pk, '_BASELINE_', a) in cascade_resolve for a in ['', 'x64', 'x86'])):
                if not pd.isna(install_dt) and not pd.isna(first_dt) and install_dt >= first_dt:
                    df.at[idx, 'Patch Evidence Status'] = 'Patch confirmed - pending rescan'
                    cascade_applied += 1
                    break

    if cascade_applied:
        log.info("Cascade resolution: %d additional rows resolved via version compliance",
                 cascade_applied)

    return df


def process_patch_match(patch_path, cve_df, min_score=9.0):
    patch  = load_data(patch_path)
    miss_p = {'Client', 'Site', 'Device', 'Status', 'Patch',
              'Discovered / Install Date'} - set(patch.columns)
    if miss_p:
        raise ValueError(f'Patch report missing required columns: {", ".join(sorted(miss_p))}')

    cve    = _drop_internal(cve_df)
    miss_c = {'Vulnerability Name', 'Name', 'Affected Products'} - set(cve.columns)
    if miss_c:
        raise ValueError(f'CVE data missing columns for patch matching: {", ".join(sorted(miss_c))}')

    if 'Customer' not in cve.columns: cve['Customer'] = ''
    if 'Site'     not in cve.columns: cve['Site']     = ''

    total_rows = len(cve)
    if 'Vulnerability Score' in cve.columns:
        cve = cve[pd.to_numeric(cve['Vulnerability Score'], errors='coerce').fillna(0) >= min_score]
    filtered_rows = len(cve)

    patch = patch.copy()
    patch['_ck']  = patch['Client'].map(_norm_compact)
    patch['_sk']  = patch['Site'].map(_norm_compact)
    patch['_dk']  = patch['Device'].map(_norm_compact)
    patch['_pk']  = patch['Patch'].map(_detect_product)
    patch['_pd']  = pd.to_datetime(patch['Discovered / Install Date'], errors='coerce')
    patch['_sr']  = patch['Status'].map(STATUS_RANK).fillna(0)
    patch['_kbs'] = patch['Patch'].apply(_extract_kbs)
    patch['_pv']  = patch['Patch'].apply(_extract_best_version)

    patch_devices = set(zip(patch['_ck'], patch['_sk'], patch['_dk']))

    cve['_ck']   = cve['Customer'].map(_norm_compact)
    cve['_sk']   = cve['Site'].map(_norm_compact)
    cve['_dk']   = cve['Name'].map(_norm_compact)
    cve['_pk']   = cve['Affected Products'].map(_detect_product)
    cve['_cves'] = cve['Vulnerability Name'].apply(_extract_cves)

    for dc in ('Date Published', 'First detected', 'Last updated'):
        if dc in cve.columns:
            cve[dc] = pd.to_datetime(
                cve[dc].astype(str).str.replace(' UTC', '', regex=False),
                errors='coerce', utc=True).dt.tz_localize(None)

    merged = cve.merge(
        patch[['_ck', '_sk', '_dk', '_pk', 'Status', 'Patch', '_pd', '_sr', '_kbs', '_pv']]
              .rename(columns={'_ck': '_mck'}),
        left_on=['_ck', '_sk', '_dk', '_pk'],
        right_on=['_mck', '_sk', '_dk', '_pk'],
        how='left', suffixes=('', '_p'),
    )
    merged = merged.sort_values(['_sr', '_pd'], ascending=[False, False], na_position='last')
    # Avoid MultiIndex.from_product overflow: keep best match per original CVE row.
    merged = merged.reset_index(drop=False)
    best   = merged.drop_duplicates(subset='index', keep='first').drop(columns='index')

    # ── Patch Match Result (vectorised) ──────────────────────────────────────
    _has_patch  = best['Patch'].notna()
    _cve_arch   = best['Affected Products'].astype(str).apply(_get_arch)
    _patch_arch = best['Patch'].fillna('').astype(str).apply(_get_arch)
    _arch_mism  = _has_patch & _cve_arch.ne('') & _patch_arch.ne('') & _cve_arch.ne(_patch_arch)
    _in_devices = pd.Series(
        [t in patch_devices for t in zip(best['_ck'], best['_sk'], best['_dk'])],
        index=best.index,
    )
    _matched_label = best['Status'].astype(str).str.strip().map(STATUS_LABEL).fillna(
        'Matched - ' + best['Status'].astype(str).str.lower()
    )
    best['Patch Match Result'] = pd.NA
    best.loc[_has_patch & ~_arch_mism,   'Patch Match Result'] = _matched_label[_has_patch & ~_arch_mism]
    best.loc[_has_patch & _arch_mism,    'Patch Match Result'] = 'Device in patch report - product not found'
    best.loc[~_has_patch & _in_devices,  'Patch Match Result'] = 'Device in patch report - product not found'
    best.loc[~_has_patch & ~_in_devices, 'Patch Match Result'] = 'Not found in patch report'

    # ── Fixed version & baseline (vectorised) ────────────────────────────────
    def _vec_fixed_version(pk_series, vname_series):
        fv_out, fs_out = [], []
        for pk, vname in zip(pk_series, vname_series):
            rules = FIXED_VERSION_RULES.get(pk, {})
            fv, fs = '', ''
            for cve in _extract_cves(str(vname)):
                if cve in rules:
                    fv, fs = rules[cve], f'config rule ({cve})'
                    break
            fv_out.append(fv); fs_out.append(fs)
        return fv_out, fs_out

    def _vec_baseline(pk_series):
        bl_out, bs_out = [], []
        for pk in pk_series:
            rules = FIXED_VERSION_RULES.get(pk, {})
            bl = rules.get('_baseline', '').strip()
            bl_out.append(bl); bs_out.append('rolling baseline' if bl else '')
        return bl_out, bs_out

    _pk_s = best.get('_pk', pd.Series([''] * len(best), index=best.index))
    _fv_vals, _fs_vals = _vec_fixed_version(_pk_s, best['Vulnerability Name'])
    best['Fixed Version Used']      = _fv_vals
    best['Fixed Version Source']    = _fs_vals
    _bl_vals, _bs_vals = _vec_baseline(_pk_s)
    best['Product Baseline']        = _bl_vals
    best['Product Baseline Source'] = _bs_vals

    best['Matched Patch Version'] = best['_pv'].fillna('')
    best['Matched KBs']           = best['_kbs'].apply(
        lambda v: ', '.join(v) if isinstance(v, list) else '')

    # ── Version Check Result (vectorised) ────────────────────────────────────
    _status_s  = best['Status'].astype(str).str.strip()
    _pv_s      = best['Matched Patch Version'].astype(str).str.strip()
    _fv_s      = best['Fixed Version Used'].astype(str).str.strip()
    _bl_s      = best['Product Baseline'].astype(str).str.strip()
    _inst_mask = _status_s.isin(INSTALLED_STATUSES)

    def _vec_vcr(status_s, pv_s, fv_s, inst_mask):
        out = []
        for status, pv, fv, inst in zip(status_s, pv_s, fv_s, inst_mask):
            if not inst:
                out.append('Patch not yet installed' if status else 'No patch evidence')
            elif not fv:
                out.append('Installed version found - no fixed baseline' if pv else 'Installed - version unknown')
            elif not pv:
                out.append('Fixed baseline known - installed version not found')
            else:
                cmp = _version_gte(pv, fv)
                out.append('Version compliant' if cmp is True else
                           'Below fixed version' if cmp is False else
                           'Version comparison failed')
        return out

    def _vec_bc(status_s, pv_s, bl_s, inst_mask):
        out = []
        for status, pv, bl, inst in zip(status_s, pv_s, bl_s, inst_mask):
            if not inst:  out.append('Not installed')
            elif not bl:  out.append('No baseline defined')
            elif not pv:  out.append('Version unknown')
            else:
                cmp = _version_gte(pv, bl)
                out.append('Compliant' if cmp is True else
                           'Below baseline' if cmp is False else
                           'Version unknown')
        return out

    best['Version Check Result'] = _vec_vcr(_status_s, _pv_s, _fv_s, _inst_mask)
    best['Baseline Compliance']  = _vec_bc(_status_s, _pv_s, _bl_s, _inst_mask)

    best = best.rename(columns={'Patch': 'Matched Patch', '_pd': 'Patch Install Date'})

    # ── Patch Evidence Status (vectorised) ───────────────────────────────────
    _vcr_s   = best['Version Check Result'].astype(str).str.strip()
    _inst_dt = pd.to_datetime(best['Patch Install Date'], errors='coerce')
    _cve_dates_max = best[['First detected', 'Date Published']].apply(
        lambda col: pd.to_datetime(
            col.astype(str).str.replace(' UTC', '', regex=False),
            errors='coerce', utc=True
        ).dt.tz_localize(None) if col.name in best.columns else pd.NaT,
        axis=0,
    ).max(axis=1)

    def _vec_pes(status_s, vcr_s, inst_dt_s, cve_max_s):
        out = []
        for status, vcr, inst_dt, cve_max in zip(status_s, vcr_s, inst_dt_s, cve_max_s):
            if status not in INSTALLED_STATUSES: out.append('Unresolved'); continue
            if vcr in ('Below fixed version', 'No fixed baseline defined', ''): out.append('Unresolved'); continue
            if 'version compliant' not in vcr.lower(): out.append('Unresolved'); continue
            if pd.isna(inst_dt) or pd.isna(cve_max): out.append('Unresolved'); continue
            out.append('Patch confirmed - pending rescan' if inst_dt >= cve_max else 'Unresolved')
        return out

    best['Patch Evidence Status'] = _vec_pes(_status_s, _vcr_s, _inst_dt, _cve_dates_max)

    best = _apply_cascade_resolution(best)

    best = best.drop(columns=[c for c in best.columns if c.startswith('_')], errors='ignore')

    ov_cols = ['Name', 'Device Type', 'Threat Status', 'Vulnerability Score',
               'Affected Products', 'Date Published', 'First detected', 'Last updated',
               'Last Response', 'Matched Patch', 'Patch Install Date',
               'Patch Match Result', 'Patch Evidence Status',
               'Product Baseline', 'Baseline Compliance']
    overview = _make_excel_safe(best[[c for c in ov_cols if c in best.columns]])
    return overview, _make_excel_safe(best), _make_excel_safe(patch), total_rows, filtered_rows


# ==============================================================================
# DATA PIPELINE: TREND / MONTH-OVER-MONTH COMPARISON
# ==============================================================================

def load_previous_report(file_path):
    try:
        xl = pd.ExcelFile(file_path, engine=_XLSX_ENGINE)
    except PermissionError:
        fname = os.path.basename(file_path)
        raise ValueError(f"'{fname}' is currently open in Excel.\n\nPlease close the file in Excel and try again.")
    except FileNotFoundError:
        fname = os.path.basename(file_path)
        raise ValueError(f"'{fname}' could not be found.\n\nPlease check the file path and try again.")
    except Exception as e:
        fname = os.path.basename(file_path)
        raise ValueError(f"Could not open '{fname}'.\n\nDetails: {e}")

    _DASHBOARD_SHEETS = {'Raw Data', 'All Detections'}
    if not (_DASHBOARD_SHEETS & set(xl.sheet_names)):
        fname = os.path.basename(file_path)
        _is_inventory = any('inventory' in s.lower() or 'device' in s.lower() for s in xl.sheet_names)
        if _is_inventory:
            _hint = ("'" + fname + "' looks like a Device Inventory / RMM report. "
                     "Select the previously generated dashboard .xlsx file instead.")
        else:
            _hint = ("'" + fname + "' does not appear to be a dashboard generated by this tool. "
                     "The Previous Dashboard field expects a .xlsx produced by a prior run, "
                     "not a raw CVE export or inventory report.")
        _sheets_str = ', '.join(xl.sheet_names[:6]) + (' ...' if len(xl.sheet_names) > 6 else '')
        raise ValueError("Cannot use '" + fname + "' as a previous dashboard.\n\n" + _hint + "\n\nSheets found: " + _sheets_str)

    target = next((s for s in ('Raw Data', 'All Detections') if s in xl.sheet_names), xl.sheet_names[0])
    df = xl.parse(target)

    missing = {'Name', 'Vulnerability Name'} - set(df.columns)
    if missing:
        raise ValueError(f"Previous report sheet '{target}' is missing columns: {', '.join(sorted(missing))}.\nPlease load a dashboard generated by this tool.")

    prev_rename = {}
    for col in df.columns:
        c = str(col).strip().lower()
        if c in ('customer name', 'client name', 'client') and 'Customer' not in df.columns:
            prev_rename[col] = 'Customer'
        elif c in ('site name', 'location name') and 'Site' not in df.columns:
            prev_rename[col] = 'Site'
        elif c == 'first detected' and col != 'First detected':
            prev_rename[col] = 'First detected'
        elif c == 'last updated' and col != 'Last updated':
            prev_rename[col] = 'Last updated'
    if prev_rename:
        df.rename(columns=prev_rename, inplace=True)

    df['_Name_Key'] = _normalize_device_col(df['Name'])
    df['_CVE_Key']  = df['Vulnerability Name'].astype(str).apply(extract_cve_id)
    df['Vulnerability Score'] = pd.to_numeric(df.get('Vulnerability Score', 0), errors='coerce').fillna(0)

    _RESERVED = {
        'trend summary', 'overview', 'all detections', 'raw data',
        'stale excluded devices', 'new this month', 'new device-cve pairs', 'new cve types',
        'resolved', 'persisting cves',
        'patch match overview', 'patch match full data', 'patch report (full)',
        'patch confirmed', 'resolved (patch confirmed)', "cves on stale devices",
        'client summary', 'summary',
    }
    resolved_pairs = set()
    # Stream product sheets with openpyxl read_only — avoids full DataFrame
    # deserialisation per sheet (xl.parse was ~2-3s on large workbooks).
    try:
        import openpyxl as _opx
        _wb_ro = _opx.load_workbook(file_path, read_only=True, data_only=True)
        for _ws in _wb_ro.worksheets:
            if _ws.title.lower() in _RESERVED:
                continue
            try:
                _header = None
                for _hr in _ws.iter_rows(min_row=1, max_row=1, values_only=True):
                    _header = [str(h).strip().lower() if h is not None else '' for h in _hr]
                    break
                if not _header or 'resolved' not in _header:
                    continue
                if 'name' not in _header or 'vulnerability name' not in _header:
                    continue
                _ri = _header.index('resolved')
                _ni = _header.index('name')
                _vi = _header.index('vulnerability name')
                _max_i = max(_ri, _ni, _vi)
                for _row in _ws.iter_rows(min_row=2, values_only=True):
                    if len(_row) > _max_i and str(_row[_ri]).strip() == '☑':
                        resolved_pairs.add((
                            normalize_device_name(str(_row[_ni] or '')),
                            extract_cve_id(str(_row[_vi] or '')),
                        ))
            except Exception:
                continue
        _wb_ro.close()
    except ImportError:
        for sheet in xl.sheet_names:
            if sheet.lower() in _RESERVED:
                continue
            try:
                sdf = xl.parse(sheet)
                if not {'Resolved', 'Name', 'Vulnerability Name'}.issubset(sdf.columns):
                    continue
                checked = sdf[sdf['Resolved'].astype(str).str.strip() == '☑']
                for nk, ck in zip(
                    checked['Name'].apply(normalize_device_name),
                    checked['Vulnerability Name'].apply(extract_cve_id),
                ):
                    resolved_pairs.add((nk, ck))
            except Exception:
                continue

    # ── Raw Data is the single source of truth ─────────────────────────────────
    # Do NOT attach _Checkbox_Resolved to df. Attaching it would let the column
    # flow into _active_trend_scope and silently hide CVEs from the Persisting set.
    # Return resolved_pairs as a standalone set so compute_trends can use it only
    # for re-detection tracking without it ever touching trend arithmetic.
    return df, resolved_pairs


def _active_trend_scope(df: pd.DataFrame, threshold: float,
                        inventory_devices=None,
                        stale_devices=None) -> pd.DataFrame:
    """
    Produce a clean, consistently-keyed DataFrame for trend arithmetic.

    Applies the full pipeline in one place so every caller uses identical logic:
      • Score threshold
      • UNRESOLVED-only (status column named 'Threat Status' or 'Status')
      • Inventory filter (decommissioned devices dropped)
      • Stale filter (stale devices dropped completely from trend comparison)
      • Deduplication on (_Name_Key, _CVE_Key, _Product_Key)
    """
    out = df.copy()
    out['_Name_Key']    = _normalize_device_col(out['Name'])
    out['_CVE_Key']     = out['Vulnerability Name'].astype(str).apply(extract_cve_id)
    out['_Product_Key'] = (
        out['Affected Products'].astype(str).apply(_detect_product)
        if 'Affected Products' in out.columns else ''
    )

    out = out[out['Vulnerability Score'] >= threshold].copy()

    _sc = ('Threat Status' if 'Threat Status' in out.columns
           else 'Status'   if 'Status'        in out.columns
           else None)
    if _sc:
        out = out[out[_sc].astype(str).str.strip().str.upper().eq('UNRESOLVED')].copy()

    if inventory_devices:
        out = out[out['_Name_Key'].isin(inventory_devices)].copy()

    if stale_devices:
        out = out[~out['_Name_Key'].isin(stale_devices)].copy()

    out = (
        out.sort_values('Vulnerability Score', ascending=False)
           .drop_duplicates(subset=['_Name_Key', '_CVE_Key', '_Product_Key'], keep='first')
    )

    if 'Base Product' not in out.columns:
        out['Base Product'] = out['Affected Products'].apply(get_base_product)

    return out


def compute_trends(current_df, previous_df, threshold,
                   inventory_devices: set = None,
                   stale_devices: set = None,
                   prev_resolved_pairs: set = None):
    """
    Compare current and previous reports at or above the score threshold.
    Raw Data is strictly honored as the single source of truth for active vulnerabilities.

    prev_resolved_pairs: set of (normalised_device, cve_id) tuples returned by
        load_previous_report alongside the DataFrame. Used ONLY for re-detection
        tracking — never to exclude rows from trend arithmetic.
    """
    cur  = current_df.copy()
    cur['_Name_Key'] = cur['Name'].apply(normalize_device_name)
    cur['_CVE_Key']  = cur['Vulnerability Name'].apply(extract_cve_id)

    prev = previous_df.copy()

    cur_t  = _active_trend_scope(current_df,  threshold, inventory_devices, stale_devices)
    prev_t = _active_trend_scope(previous_df, threshold, inventory_devices, stale_devices)

    if inventory_devices or stale_devices:
        dropped = len(_active_trend_scope(previous_df, threshold)) - len(prev_t)
        if dropped > 0:
            log.info("Trend: excluded %d previous-period row(s) for decommissioned or stale devices", dropped)

    def _kev_count(df):
        if 'CISA KEV' not in df.columns: return 0
        return int(df[df['CISA KEV'].astype(str).str.strip().str.lower()
                       .isin(['yes', 'true', '1', 'y'])]['Vulnerability Name'].nunique())

    def _exploit_count(df):
        if 'Has Known Exploit' not in df.columns: return 0
        return int(df[df['Has Known Exploit'].astype(str).str.strip().str.lower()
                       .isin(['yes', 'true', '1', 'y'])]['Vulnerability Name'].nunique())

    def _srv_count(df):
        if 'Device Type' not in df.columns: return 0
        return int(df[df['Device Type'] == 'Server']['Name'].nunique())

    snap_prev_cves    = prev_t['_CVE_Key'].nunique()
    snap_cur_cves     = cur_t['_CVE_Key'].nunique()
    snap_prev_devices = int(prev_t['Name'].nunique())
    snap_cur_devices  = int(cur_t['Name'].nunique())
    snap_prev_kev     = _kev_count(prev_t)
    snap_cur_kev      = _kev_count(cur_t)
    snap_prev_exploit = _exploit_count(prev_t)
    snap_cur_exploit  = _exploit_count(cur_t)
    snap_prev_servers = _srv_count(prev_t)
    snap_cur_servers  = _srv_count(cur_t)

    common_products = (set(cur_t['Base Product'].unique()) & set(prev_t['Base Product'].unique()))
    new_products    = set(cur_t['Base Product'].unique()) - set(prev_t['Base Product'].unique())

    cur_scoped  = cur_t[cur_t['Base Product'].isin(common_products)].copy()
    prev_scoped = prev_t[prev_t['Base Product'].isin(common_products)].copy()

    if new_products:
        log.info("Trend: %d product(s) new this period (not in previous report): %s",
                 len(new_products), sorted(new_products))

    # ── Re-detection tracking (checkbox triage data ONLY) ─────────────────────
    # prev_resolved_pairs arrives directly from load_previous_report's second
    # return value so it never touches the DataFrames feeding _active_trend_scope
    # and cannot influence New / Resolved / Persisting counts.
    checkbox_resolved   = prev_resolved_pairs or set()
    cur_active_pairs_2d = set(zip(cur_t['_Name_Key'], cur_t['_CVE_Key']))
    redetected_pairs    = checkbox_resolved & cur_active_pairs_2d
    redetected_count    = len(redetected_pairs)

    # ── Set Arithmetic ──
    # NOTE: We strictly DO NOT filter cur_scoped based on previous manual checkboxes. 
    # Raw data wins. If it is UNRESOLVED in the raw data, it persists.
    cur_all_pair_keys  = set(zip(cur_t['_Name_Key'], cur_t['_CVE_Key'], cur_t['_Product_Key']))
    cur_scoped_pair_keys = set(zip(cur_scoped['_Name_Key'], cur_scoped['_CVE_Key'], cur_scoped['_Product_Key']))
    prev_pair_keys     = set(zip(prev_scoped['_Name_Key'], prev_scoped['_CVE_Key'], prev_scoped['_Product_Key']))

    new_pair_keys        = cur_all_pair_keys  - prev_pair_keys
    resolved_pair_keys   = prev_pair_keys     - cur_scoped_pair_keys
    persisting_pair_keys = cur_scoped_pair_keys & prev_pair_keys

    def _filter_pairs(df, keys):
        if not keys or df.empty:
            return _drop_internal(df.iloc[0:0].copy())
        _key_df = pd.DataFrame(list(keys), columns=['_Name_Key', '_CVE_Key', '_Product_Key'])
        return _drop_internal(df.merge(_key_df, on=['_Name_Key', '_CVE_Key', '_Product_Key'], how='inner').copy())

    new_pairs_df      = _filter_pairs(cur_t,       new_pair_keys).sort_values('Vulnerability Score', ascending=False)
    resolved_df       = _filter_pairs(prev_scoped,  resolved_pair_keys).sort_values('Vulnerability Score', ascending=False)
    persisting_df     = _filter_pairs(cur_scoped,   persisting_pair_keys).sort_values('Vulnerability Score', ascending=False)

    # CVE-type sets
    cur_all_cve_ids  = set(cur_t['_CVE_Key'].unique())
    cur_cve_ids      = set(cur_scoped['_CVE_Key'].unique())
    prev_cve_ids     = set(prev_scoped['_CVE_Key'].unique())

    new_cve_ids          = cur_all_cve_ids - prev_cve_ids
    resolved_cve_ids     = prev_cve_ids - cur_cve_ids
    persisting_cve_ids   = cur_cve_ids & prev_cve_ids

    new_cve_types_df = _drop_internal(
        cur_t[cur_t['_CVE_Key'].isin(new_cve_ids)].copy()
    ).sort_values('Vulnerability Score', ascending=False)

    cur_dev_set  = set(cur_t['_Name_Key'].unique())
    prev_dev_set = set(prev_t['_Name_Key'].unique())

    scoped_cur_cves  = len(cur_all_cve_ids)
    scoped_prev_cves = len(prev_cve_ids)

    metrics = {
        'cur_cves':             snap_cur_cves,
        'prev_cves':            snap_prev_cves,
        'cur_devices':          snap_cur_devices,
        'prev_devices':         snap_prev_devices,
        'cur_kev':              snap_cur_kev,
        'prev_kev':             snap_prev_kev,
        'cur_exploit':          snap_cur_exploit,
        'prev_exploit':         snap_prev_exploit,
        'cur_servers':          snap_cur_servers,
        'prev_servers':         snap_prev_servers,
        'scoped_cur_cves':      scoped_cur_cves,
        'scoped_prev_cves':     scoped_prev_cves,
        'new_cve_count':        len(new_cve_ids),
        'resolved_cve_count':   len(resolved_cve_ids),
        'persisting_cve_count': len(persisting_cve_ids),
        'new_pair_count':       len(new_pair_keys),
        'resolved_pair_count':  len(resolved_pair_keys),
        'persisting_pair_count':len(persisting_pair_keys),
        'new_devices':          len(cur_dev_set  - prev_dev_set),
        'remediated_devices':   len(prev_dev_set - cur_dev_set),
    }

    cur_prod      = cur_t.groupby('Base Product')['_Name_Key'].nunique()
    cur_cve_prod  = cur_t.groupby('Base Product')['_CVE_Key'].nunique()
    prev_prod     = prev_t.groupby('Base Product')['_Name_Key'].nunique()
    prev_cve_prod = prev_t.groupby('Base Product')['_CVE_Key'].nunique()
    product_trend = (
        pd.DataFrame({
            'Current':     cur_prod,
            'Previous':    prev_prod,
            'CVE_Current': cur_cve_prod,
            'CVE_Previous':prev_cve_prod,
        })
        .fillna(0).astype(int)
    )
    product_trend['Change']     = product_trend['Current']     - product_trend['Previous']
    product_trend['CVE_Change'] = product_trend['CVE_Current'] - product_trend['CVE_Previous']
    product_trend = product_trend.sort_values('Current', ascending=False).head(10)

    return {
        'metrics':                 metrics,
        'new_df':                  new_pairs_df,
        'new_pairs_df':            new_pairs_df,
        'new_cve_types_df':        new_cve_types_df,
        'resolved_df':             resolved_df,
        'persisting_df':           persisting_df,
        'product_trend':           product_trend,
        'redetected_count':        redetected_count,
    }


# ==============================================================================
# PATCH DIAGNOSTICS  (lag · version drift · mismatch)
# ==============================================================================

def compute_patch_diagnostics(patch_full_df: pd.DataFrame) -> dict:
    df = patch_full_df.copy()
    required = {'Name', 'Vulnerability Name', 'Patch Match Result',
                'Patch Evidence Status'}
    if not required.issubset(df.columns):
        log.warning("compute_patch_diagnostics: missing columns %s — skipping",
                    required - set(df.columns))
        return {'patch_lag_df': pd.DataFrame(),
                'version_drift_df': pd.DataFrame(),
                'mismatch_summary_df': pd.DataFrame()}

    lag_rows = []
    if 'Patch Install Date' in df.columns and 'First detected' in df.columns:
        for _, row in df.iterrows():
            install_dt = pd.to_datetime(row.get('Patch Install Date'), errors='coerce')
            first_dt   = pd.to_datetime(row.get('First detected'),    errors='coerce')
            if pd.isna(install_dt) or pd.isna(first_dt):
                continue
            lag_days = (install_dt - first_dt).days
            lag_rows.append({
                'Device':             row.get('Name', ''),
                'CVE':                extract_cve_id(str(row.get('Vulnerability Name', ''))),
                'Product':            row.get('Affected Products', ''),
                'First Detected':     first_dt.date(),
                'Patch Install Date': install_dt.date(),
                'Lag (days)':         lag_days,
                'Status':             row.get('Patch Evidence Status', ''),
            })
    patch_lag_df = (pd.DataFrame(lag_rows)
                    .sort_values('Lag (days)', ascending=False)
                    .reset_index(drop=True)
                    if lag_rows else pd.DataFrame())

    drift_rows = []
    if 'Matched Patch Version' in df.columns:
        df['_bp'] = df['Affected Products'].apply(get_base_product)
        for product, grp in df.groupby('_bp'):
            versions = (grp['Matched Patch Version']
                        .dropna()
                        .astype(str)
                        .str.strip()
                        .loc[lambda s: s.str.len() > 0]
                        .unique()
                        .tolist())
            if len(versions) < 2:
                continue
            parsed = [v for v in (_parse_version(v) for v in versions) if v]
            if len(parsed) < 2:
                continue
            spread = len(set(versions))
            drift_rows.append({
                'Product':          product,
                'Distinct Versions': spread,
                'Min Version':      min(versions, key=lambda v: _parse_version(v) or (0,)),
                'Max Version':      max(versions, key=lambda v: _parse_version(v) or (0,)),
                'Versions Seen':    ', '.join(sorted(set(versions))),
                'Device Count':     grp['Name'].nunique(),
            })
    version_drift_df = (pd.DataFrame(drift_rows)
                        .sort_values('Distinct Versions', ascending=False)
                        .reset_index(drop=True)
                        if drift_rows else pd.DataFrame())

    mismatch_rows = []
    for _, row in df.iterrows():
        gap = classify_patch_gap(
            row.get('Patch Match Result', ''),
            row.get('Patch Evidence Status', ''),
        )
        if gap != 'detection_mismatch':
            continue
        install_dt = pd.to_datetime(row.get('Patch Install Date'), errors='coerce')
        first_dt   = pd.to_datetime(row.get('First detected'),    errors='coerce')
        lag = (install_dt - first_dt).days if not (pd.isna(install_dt) or pd.isna(first_dt)) else None
        mismatch_rows.append({
            'Device':               row.get('Name', ''),
            'CVE':                  extract_cve_id(str(row.get('Vulnerability Name', ''))),
            'Product':              row.get('Affected Products', ''),
            'Patch Match Result':   row.get('Patch Match Result', ''),
            'Installed Version':    row.get('Matched Patch Version', ''),
            'Fixed Version Needed': row.get('Fixed Version Used', ''),
            'Patch Install Date':   row.get('Patch Install Date', ''),
            'First Detected':       row.get('First detected', ''),
            'Lag (days)':           lag,
            'Likely Cause':         (
                'Install predates CVE detection — patch may not address this CVE'
                if lag is not None and lag < 0
                else 'Patch installed but CVE still detected — scanner/patch tool disagreement'
            ),
        })
    mismatch_summary_df = (pd.DataFrame(mismatch_rows)
                           .reset_index(drop=True)
                           if mismatch_rows else pd.DataFrame())

    return {
        'patch_lag_df':       patch_lag_df,
        'version_drift_df':   version_drift_df,
        'mismatch_summary_df': mismatch_summary_df,
    }


# ==============================================================================
# PATCH FAILURE REPORT
# ==============================================================================

_FAIL_CATEGORY_MAP = {
    'reboot_pending':          'Reboot required before patch can install',
    'catalog_miss':            'Patch not found in WUA catalog — may be superseded',
    'network_timeout':         'Network timeout during patch download',
    'cert_failure':            'Certificate verification failed — PME cache may need clearing',
    'checksum_failure':        'Patch file checksum error — corrupted download',
    'feature_update_conflict': 'Feature Update in progress — retry after update completes',
    'third_party_unknown':     'Third-party patch application not recognised by RMM',
    'agent_timeout':           'RMM agent timed out during install',
    'install_error':           'Installer returned error code',
    'unknown':                 'Unknown failure',
}

def _classify_failure_reason(reason: str) -> str:
    r = str(reason).lower()
    if 'reboot is required'         in r: return 'reboot_pending'
    if 'not found in wua catalog'   in r: return 'catalog_miss'
    if 'certificate verification'   in r: return 'cert_failure'
    if 'checksum'                   in r: return 'checksum_failure'
    if 'feature update'             in r: return 'feature_update_conflict'
    if 'unknown application'        in r: return 'third_party_unknown'
    if 'timeout'                    in r: return 'network_timeout'
    if "couldn't download"          in r: return 'network_timeout'
    if 'incomplete download'        in r: return 'network_timeout'
    if 'timed out'                  in r: return 'agent_timeout'
    if 'operation was canceled'     in r: return 'agent_timeout'
    if 'value does not fall'        in r: return 'install_error'
    if 'installation error'         in r: return 'install_error'
    if 'fatal error'                in r: return 'install_error'
    if 'process exited'             in r: return 'install_error'
    return 'unknown'


def load_patch_failure_report(file_path: str) -> pd.DataFrame:
    df = load_data(file_path)

    rename = {}
    for col in df.columns:
        cl = col.lower().strip()
        if cl == 'device':         rename[col] = 'Device'
        elif cl == 'site':         rename[col] = 'Site'
        elif cl == 'client':       rename[col] = 'Client'
        elif cl == 'patch':        rename[col] = 'Patch'
        elif 'failure status' in cl: rename[col] = 'Failure Status'
        elif 'failure reason' in cl: rename[col] = 'Failure Reason'
        elif 'time' in cl:         rename[col] = 'Time'
    df = df.rename(columns=rename)

    df['_device_norm']   = df['Device'].apply(normalize_device_name)
    df['_failure_cat']   = df['Failure Reason'].apply(_classify_failure_reason)
    df['_failure_desc']  = df['_failure_cat'].map(_FAIL_CATEGORY_MAP)
    df['_kbs']           = df['Patch'].astype(str).apply(_extract_kbs)

    return df


def build_patch_failure_lookup(failure_df: pd.DataFrame) -> dict:
    result = {}
    for device, grp in failure_df.groupby('_device_norm'):
        cats    = grp['_failure_cat'].value_counts().to_dict()
        top_cat = grp['_failure_cat'].value_counts().index[0]
        result[device] = {
            'failure_count':    len(grp),
            'unique_kbs':       len({kb for kbs in grp['_kbs'] for kb in kbs}),
            'top_category':     top_cat,
            'top_description':  _FAIL_CATEGORY_MAP.get(top_cat, top_cat),
            'categories':       cats,
        }
    return result

# ==============================================================================
# BROWSER AUDIT INTEGRATION
# ==============================================================================

def load_browser_audit(file_path: str) -> pd.DataFrame:
    """
    Parse a Browser Audit xlsx (from PS browser scan script).
    Returns normalised rows with Device, Browser, Version, Architecture,
    Install Scope, Is 32-bit, Is Per-User/AppData, Username, Last Response.
    """
    try:
        xl = pd.ExcelFile(file_path, engine=_XLSX_ENGINE)
    except Exception as e:
        log.warning("load_browser_audit: cannot open %s — %s", file_path, e)
        return pd.DataFrame()

    _BROWSER_SHEETS = {
        'chrome':  'Google Chrome',
        'edge':    'Microsoft Edge',
        'firefox': 'Mozilla Firefox',
    }

    frames = []
    for sheet in xl.sheet_names:
        sl = sheet.lower()
        browser = next((v for k, v in _BROWSER_SHEETS.items() if k in sl), None)
        if browser is None:
            continue
        try:
            raw = xl.parse(sheet, header=None)
        except Exception:
            continue
        if raw.empty:
            continue

        # Detect header row (first row with ≥3 non-null values)
        header_row = None
        for idx, row in raw.iterrows():
            if len(row.dropna()) >= 3:
                header_row = idx
                break
        if header_row is None:
            continue

        df = xl.parse(sheet, header=header_row)
        df.columns = [str(c).strip() for c in df.columns]

        rename = {}
        for col in df.columns:
            cl = col.lower().strip()
            if cl == 'device':                      rename[col] = 'Device'
            elif cl == 'version':                   rename[col] = 'Version'
            elif cl == 'architecture':              rename[col] = 'Architecture'
            elif cl == 'install scope':             rename[col] = 'Install Scope'
            elif cl == 'install path':              rename[col] = 'Install Path'
            elif 'is 32' in cl:                     rename[col] = 'Is 32-bit'
            elif 'per-user' in cl or 'appdata' in cl: rename[col] = 'Is Per-User/AppData'
            elif cl == 'username':                  rename[col] = 'Username'
            elif 'last response' in cl:             rename[col] = 'Last Response'
            elif 'last used' in cl:                 rename[col] = 'Last Used'
        df = df.rename(columns=rename)

        # Skip sheets that lack required columns (e.g. Firefox Last Used)
        if 'Version' not in df.columns or 'Device' not in df.columns:
            continue

        df = df.dropna(subset=['Device', 'Version'], how='any')
        df = df[df['Version'].astype(str).str.match(r'\d+[\d.]+')]
        df['Browser'] = browser
        df['Source']  = 'Browser Audit'
        frames.append(df)

    if not frames:
        log.warning("load_browser_audit: no browser sheets found in %s", file_path)
        return pd.DataFrame()

    out = pd.concat(frames, ignore_index=True)
    out['Device']  = out['Device'].astype(str).str.strip()
    out['Version'] = out['Version'].astype(str).str.strip()
    return out


def merge_browser_audit_into_drift(version_drift_df: pd.DataFrame,
                                   browser_audit_df: pd.DataFrame) -> pd.DataFrame:
    """
    Enrich version_drift_df with per-user/AppData and 32-bit install counts
    from the browser audit. Adds an 'Audit Note' column.
    """
    if browser_audit_df is None or browser_audit_df.empty:
        return version_drift_df

    out = version_drift_df.copy() if not version_drift_df.empty else pd.DataFrame()

    _BA_LABEL = {
        'Google Chrome':   ['Google Chrome'],
        'Microsoft Edge':  ['Microsoft Edge'],
        'Mozilla Firefox': ['Mozilla Firefox'],
    }

    audit_rows = []
    for browser, labels in _BA_LABEL.items():
        grp = browser_audit_df[browser_audit_df['Browser'].isin(labels)]
        if grp.empty:
            continue

        versions = (grp['Version'].dropna().astype(str).str.strip()
                    .pipe(lambda s: s[s.str.match(r'\d+[\d.]+')]))
        uniq_vers = sorted(versions.unique().tolist(),
                           key=lambda v: tuple(int(x) for x in re.findall(r'\d+', v)) if re.findall(r'\d+', v) else (0,))
        n_devices = grp['Device'].nunique()

        pu_count  = int(grp['Is Per-User/AppData'].astype(str).str.strip().str.lower()
                        .isin(['yes', 'true', '1']).sum()) if 'Is Per-User/AppData' in grp.columns else 0
        bit32_cnt = int(grp['Is 32-bit'].astype(str).str.strip().str.lower()
                        .isin(['yes', 'true', '1']).sum()) if 'Is 32-bit' in grp.columns else 0

        notes = []
        if pu_count:   notes.append(f'{pu_count} per-user/AppData install(s)')
        if bit32_cnt:  notes.append(f'{bit32_cnt} 32-bit install(s) on 64-bit OS')
        note_str = '; '.join(notes) if notes else '—'

        audit_rows.append({
            'Product':           browser,
            'Distinct Versions': len(uniq_vers),
            'Min Version':       uniq_vers[0]  if uniq_vers else '',
            'Max Version':       uniq_vers[-1] if uniq_vers else '',
            'Versions Seen':     ', '.join(uniq_vers),
            'Device Count':      n_devices,
            'Audit Note':        note_str,
        })

    if not audit_rows:
        return out

    audit_df = pd.DataFrame(audit_rows)

    if out.empty:
        out = audit_df.copy()
        return out

    if 'Audit Note' not in out.columns:
        out['Audit Note'] = '—'

    for _, arow in audit_df.iterrows():
        prod = arow['Product']
        mask = out['Product'].astype(str).str.startswith(prod)
        if mask.any():
            out.loc[mask, 'Audit Note'] = arow['Audit Note']
        else:
            out = pd.concat([out, pd.DataFrame([arow.to_dict()])], ignore_index=True)

    return out