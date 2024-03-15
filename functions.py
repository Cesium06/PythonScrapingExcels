import re


def get_loop_info(bus_route_color):
    # Regular expression pattern to match "Blue Loop" and capture it
    pattern = r"([A-Za-z]+)\sLoop"

    result = ''

    # Using re.match to find the pattern at the beginning of the string
    match = re.match(pattern, bus_route_color)

    if not match:
        pattern = r"([A-Za-z]+)\sLink"
        match = re.match(pattern, bus_route_color)

    # Extracting the captured group if a match is found
    if match:
        result = match.group(1)

    return result


def get_day_type(bus_route_color):
    # Regular expression pattern to match and capture the part after the colon
    pattern = r'^[A-Za-z ]+: (.*)$'
    result = ''

    # Using re.match to find the pattern at the beginning of the string
    match = re.match(pattern, bus_route_color)

    # Extracting the captured group if a match is found
    if match:
        result = match.group(1)

    if result == 'Mon - Fri':
        return 'weekdays'
    elif result == 'Saturday':
        return 'saturdays'
    elif result == 'Sunday':
        return 'sunday'
    else:
        return 'N', 'N', 'N'


def month_start_end(date_range):
    # Regular expression to match month, start day, and end day
    pattern = r"([A-Za-z]{3,4})\s(\d{1,2})\s*-\s*(\d{1,2})"

    matches = re.search(pattern, date_range.iloc[0])
    if matches:
        month = matches.group(1)
        start_day = matches.group(2)
        end_day = matches.group(3)
        days = date_range['Unnamed: 3']

    return month, start_day, end_day, days
