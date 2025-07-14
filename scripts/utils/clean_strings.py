def format_raw_value(val) -> str:
    if val is None:
        return ""

    # Ugly nan-check, but avoids additional conversion
    match str(val):
        case "nan":
            return "—" # Note: em-dash
        case "0":
            return ""
        case res:
            return res