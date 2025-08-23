def format_number(value):
    """Format a number with commas, 2 decimals, and parentheses for negatives."""
    try:
        abs_value = abs(value)
        formatted = f"{abs_value:,.2f}" if abs_value >= 1000 else f"{abs_value:.2f}"
        return f"({formatted})" if value < 0 else formatted
    except Exception:
        return str(value)

