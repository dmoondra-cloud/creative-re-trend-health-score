"""
Configuration for Creative RE Underwriting Suite
Customize categorization rules, templates, and behavior here
"""

# ============================================================================
# CATEGORIZATION RULES
# ============================================================================

# Custom categorization patterns (extends built-in rules)
CUSTOM_CATEGORIES = {
    'Example Custom Income': {
        'patterns': [r'Custom Pattern 1', r'Custom Pattern 2'],
        'type': 'income',
        'section': 'Other Income'
    },
    # Add more custom rules here
}

# Properties with special handling
PROPERTY_OVERRIDES = {
    'Woodward Park': {
        'gpr_calculation': 'special_method',  # Custom GPR calculation
        'occupancy_method': 'custom'
    }
}

# ============================================================================
# TEMPLATE PATHS
# ============================================================================

# Location of default templates (relative to app directory)
DEFAULT_TEMPLATE_PATH = 'templates/THS_Template_Default.xlsx'

# Template sheet names
TEMPLATE_SHEETS = {
    'THS_MAIN': 'Trend Health Score',
    'THS_ANCHORS': 'Anchors & Notes',
    'T12': 'T12',
    'RR': 'Rent Roll'
}

# ============================================================================
# PARSER SETTINGS
# ============================================================================

# Month detection settings
MIN_MONTHS_REQUIRED = 3
MAX_MONTHS_EXPECTED = 12

# Property name extraction - search scope
PROPERTY_SEARCH_ROWS = (1, 15)

# Date extraction - search scope
DATE_SEARCH_ROWS = (1, 12)

# ============================================================================
# CATEGORIZATION THRESHOLDS
# ============================================================================

# Confidence thresholds for user warnings
LOW_CONFIDENCE_THRESHOLD = 0.6
WARNING_CONFIDENCE_THRESHOLD = 0.7

# Items requiring manual review if confidence < threshold
REQUIRE_REVIEW = True

# ============================================================================
# VALIDATION SETTINGS
# ============================================================================

# Enable/disable specific validations
VALIDATIONS = {
    'preserve_formulas': True,
    'check_data_types': True,
    'validate_gpr': True,
    'validate_occupancy': True,
}

# Acceptable value ranges for validations
VALUE_RANGES = {
    'occupancy_pct': (0, 100),
    'concessions_pct': (-100, 100),
    'bad_debt_pct': (0, 100),
}

# ============================================================================
# UI SETTINGS
# ============================================================================

# Streamlit page config
APP_TITLE = "Creative RE Underwriting Suite"
APP_ICON = "📊"
APP_LAYOUT = "wide"

# Display settings
SHOW_CONFIDENCE_SCORES = True
SHOW_VALIDATION_WARNINGS = True
SHOW_DEBUG_INFO = False

# ============================================================================
# PROCESSING SETTINGS
# ============================================================================

# Temporary file settings
TEMP_DIR = 'temp/'
KEEP_TEMP_FILES = False  # Set True for debugging

# Cache settings
ENABLE_CACHE = True
CACHE_EXPIRY_MINUTES = 60

# ============================================================================
# ADVANCED FEATURES
# ============================================================================

# Enable beta features
BETA_FEATURES = {
    'batch_processing': False,
    'api_endpoint': False,
    'custom_template_builder': False,
    'historical_tracking': False,
}

# Special handling for specific line items
SPECIAL_HANDLERS = {
    'RUBS': {
        'negate_if_negative': True,
        'category': 'Other Income'
    },
    'Utility Income': {
        'negate_if_negative': True,
        'category': 'Other Income'
    },
    'Reimbursement': {
        'always_income': True,
        'category': 'Other Income'
    },
}
