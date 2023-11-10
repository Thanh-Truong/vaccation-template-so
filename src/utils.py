import re
def parse_input_periods(custom_periods):
    # Extract substrings matching the pattern 
    # digit-space*-space*hyphen-space*\d+
    months_per_sheet = re.findall(r'\d+\s*-\s*\d+', custom_periods)

    # Split, then strip(), then make integers
    return  [tuple(map(int, map(str.strip, 
                        period.split('-')))) for period in months_per_sheet]