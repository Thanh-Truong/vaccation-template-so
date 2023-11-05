from datetime import date
import requests
import date_utils
import json

def cache_calendar(year):
    # Save the dictionary to a JSON file
    url = f"https://api.dryg.net/dagar/v2.1/{year}"
    response = requests.get(url)
    data = response.json()

    with open(f"calendar/{year}.json", "w") as json_file:
        json.dump(data, json_file)

def load_calendar(year):
    file_path = f"calendar/{year}.json"
    try:
        with open(file_path, "r") as json_file:
            return json.load(json_file)      
    except FileNotFoundError:
        # Not cached yet 
        cache_calendar(year)
        return load_calendar(year)

def get_swedish_holidays(year):
    calendar = load_calendar(year)

    if "dagar" in calendar:
        holidays = [(day["datum"], day["helgdag"]) 
                    for day in calendar["dagar"] 
                        if day["röd dag"]=='Ja' 
                            and "helgdag" in day]
        return holidays
    return []

def get_swedish_holidays_as_date_description(year):
    return [(date_utils.parse_date(holiday_str), holiday_description) 
                for holiday_str, holiday_description 
                    in get_swedish_holidays(year)]

def get_holiday_description(date_obj, holidays):
    for day, _ in holidays:
        if day == date_obj:
            return _
    return None

if __name__ == "__main__":
    for year in range(2024, 2028):
        cache_calendar(year)
    
    for day, descr in get_swedish_holidays(2028):
        print(f"{day} - {descr}")
