import json
from copy import deepcopy
from pathlib import Path


CONFIG_DIR = Path.home() / ".caseware_ocr"
CONFIG_PATH = CONFIG_DIR / "config.json"
VENDOR_MAP_PATH = CONFIG_DIR / "vendor_map.json"
ACCOUNTS_MAP_PATH = CONFIG_DIR / "accounts_map.json"

DEFAULT_PROFILE_NAME = "default"
DEFAULT_MODEL = "nvidia/nemotron-nano-12b-v2-vl:free"
DEFAULT_EXPORT_FORMAT = "caseware"

DEFAULT_PROFILE = {
    "api_key": "",
    "model": DEFAULT_MODEL,
    "currency": "CAD",
    "max_pages": 12,
    "rpm": 12,
    "zoom": 2.0,
    "max_retries": 5,
    "retry_backoff": 5,
    "retry_max_sleep": 60,
    "refine_amounts": True,
    "recursive": True,
    "use_cache": True,
    "export_format": DEFAULT_EXPORT_FORMAT,
}

DEFAULT_CONFIG = {
    "active_profile": DEFAULT_PROFILE_NAME,
    "profiles": {
        DEFAULT_PROFILE_NAME: deepcopy(DEFAULT_PROFILE),
    },
}

DEFAULT_VENDOR_MAP = {
    "AMZN MKTP": "Amazon",
    "AMAZON.CA": "Amazon",
    "AMAZON": "Amazon",
    "COSTCO WHOLESALE": "Costco",
    "COSTCO": "Costco",
    "TIM HORTONS": "Tim Hortons",
    "TIMS": "Tim Hortons",
    "STARBUCKS": "Starbucks",
    "WALMART": "Walmart",
    "HOME DEPOT": "Home Depot",
    "UBER *EATS": "Uber Eats",
    "UBER": "Uber",
    "SHELL": "Shell",
    "ESSO": "Esso",
    "PETRO-CANADA": "Petro-Canada",
    "ROGERS": "Rogers",
    "BELL CANADA": "Bell",
    "TELUS": "Telus",
    "CANADA POST": "Canada Post",
}

DEFAULT_ACCOUNTS_MAP = {
    "Amazon": "Office Supplies",
    "Costco": "Office Supplies",
    "Tim Hortons": "Meals and Entertainment",
    "Starbucks": "Meals and Entertainment",
    "Walmart": "Office Supplies",
    "Home Depot": "Repairs and Maintenance",
    "Uber Eats": "Meals and Entertainment",
    "Uber": "Travel",
    "Shell": "Vehicle Expenses",
    "Esso": "Vehicle Expenses",
    "Petro-Canada": "Vehicle Expenses",
    "Rogers": "Telephone and Internet",
    "Bell": "Telephone and Internet",
    "Telus": "Telephone and Internet",
    "Canada Post": "Postage and Delivery",
}


def _ensure_config_dir():
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)


def _read_json(path: Path, default_value):
    _ensure_config_dir()
    if not path.exists():
        _write_json(path, default_value)
        return deepcopy(default_value)
    try:
        with path.open("r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError):
        _write_json(path, default_value)
        return deepcopy(default_value)


def _write_json(path: Path, data):
    _ensure_config_dir()
    with path.open("w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def _merge_profile(profile_data):
    merged = deepcopy(DEFAULT_PROFILE)
    if isinstance(profile_data, dict):
        merged.update(profile_data)
    return merged


def load_config():
    cfg = _read_json(CONFIG_PATH, DEFAULT_CONFIG)
    if not isinstance(cfg, dict):
        cfg = deepcopy(DEFAULT_CONFIG)

    profiles = cfg.get("profiles")
    if not isinstance(profiles, dict) or not profiles:
        profiles = {DEFAULT_PROFILE_NAME: deepcopy(DEFAULT_PROFILE)}
    else:
        profiles = {name: _merge_profile(data) for name, data in profiles.items()}

    active_profile = cfg.get("active_profile")
    if not isinstance(active_profile, str) or active_profile not in profiles:
        active_profile = DEFAULT_PROFILE_NAME
        if active_profile not in profiles:
            profiles[active_profile] = deepcopy(DEFAULT_PROFILE)

    normalized = {"active_profile": active_profile, "profiles": profiles}
    save_config(normalized)
    return normalized


def save_config(cfg):
    if not isinstance(cfg, dict):
        cfg = deepcopy(DEFAULT_CONFIG)
    profiles = cfg.get("profiles", {})
    if not isinstance(profiles, dict):
        profiles = {}
    if not profiles:
        profiles[DEFAULT_PROFILE_NAME] = deepcopy(DEFAULT_PROFILE)
    normalized_profiles = {name: _merge_profile(data) for name, data in profiles.items()}
    active_profile = cfg.get("active_profile")
    if active_profile not in normalized_profiles:
        active_profile = next(iter(normalized_profiles.keys()))
    normalized = {"active_profile": active_profile, "profiles": normalized_profiles}
    _write_json(CONFIG_PATH, normalized)
    return normalized


def get_profile(cfg, name):
    if not isinstance(cfg, dict):
        cfg = load_config()
    profiles = cfg.get("profiles", {})
    if not isinstance(profiles, dict):
        profiles = {}
    if name in profiles:
        return _merge_profile(profiles[name])
    return _merge_profile(None)


def save_profile(cfg, name, data):
    if not isinstance(cfg, dict):
        cfg = load_config()
    if not isinstance(name, str) or not name.strip():
        name = DEFAULT_PROFILE_NAME
    profiles = cfg.get("profiles")
    if not isinstance(profiles, dict):
        profiles = {}
    profiles[name] = _merge_profile(data)
    cfg["profiles"] = profiles
    cfg["active_profile"] = name
    return save_config(cfg)


def load_vendor_map():
    data = _read_json(VENDOR_MAP_PATH, DEFAULT_VENDOR_MAP)
    if isinstance(data, dict):
        return data
    save_vendor_map(DEFAULT_VENDOR_MAP)
    return deepcopy(DEFAULT_VENDOR_MAP)


def save_vendor_map(data):
    if not isinstance(data, dict):
        data = deepcopy(DEFAULT_VENDOR_MAP)
    _write_json(VENDOR_MAP_PATH, data)
    return data


def load_accounts_map():
    data = _read_json(ACCOUNTS_MAP_PATH, DEFAULT_ACCOUNTS_MAP)
    if isinstance(data, dict):
        return data
    save_accounts_map(DEFAULT_ACCOUNTS_MAP)
    return deepcopy(DEFAULT_ACCOUNTS_MAP)


def save_accounts_map(data):
    if not isinstance(data, dict):
        data = deepcopy(DEFAULT_ACCOUNTS_MAP)
    _write_json(ACCOUNTS_MAP_PATH, data)
    return data

