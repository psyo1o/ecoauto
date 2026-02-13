def normalize_data(data):
    """
    Normalize the data to a common scale.
    Raises: ValueError if data is not a list or if it's empty.
    """
    if not isinstance(data, list) or len(data) == 0:
        raise ValueError("Input must be a non-empty list.")

    min_val = min(data)
    max_val = max(data)
    range_val = max_val - min_val

    if range_val == 0:
        return [0.0 for _ in data]

    normalized = [(x - min_val) / range_val for x in data]
    return normalized


def parse_data(data_str):
    """
    Parse a comma-separated string of data into a list of floats.
    Raises: ValueError if the input is empty or not valid.
    """
    if not data_str:
        raise ValueError("Input string cannot be empty.")

    try:
        data_list = [float(value) for value in data_str.split(",")]
    except ValueError:
        raise ValueError("Invalid value found in input string.")

    return data_list
