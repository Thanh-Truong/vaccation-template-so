from datetime import date
import holidays
import requests

def get_swedish_public_holidays(year):
    # Create a dictionary of Swedish holidays for the specified year
    swedish_holidays = holidays.Sweden(years=year)

    # Filter out Sundays from the list of public holidays
    public_holidays_filtered = [(d, description) for d, description in swedish_holidays.items() if d.weekday() != 6]

    return public_holidays_filtered

def get_swedish_holidays(year):
    url = f"https://api.dryg.net/dagar/v2.1/{year}"
    response = requests.get(url)
    data = response.json()

    if "dagar" in data:
        holidays = [(day["datum"], day["helgdag"]) 
                    for day in data["dagar"] 
                        if day["röd dag"]=='Ja' 
                            and "helgdag" in day]
        return holidays
    return []

if __name__ == "__main__":
    # Example usage:
    year = 2024  # Replace with the year you're interested in
    public_holidays_2024 = get_swedish_public_holidays(year)
    for date, description in public_holidays_2024:
       print(date, " - ", description)

    # Användning
    year = 2024  # Byt ut med det år du är intresserad av
    swedish_holidays_2024 = get_swedish_holidays(year)
    for date, description in swedish_holidays_2024:
       print(date, " - ", description)