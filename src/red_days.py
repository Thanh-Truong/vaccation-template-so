from datetime import date
import requests
import date_utils as date_utils
import json
from openpyxl.drawing.image import Image

def cache_calendar(year):
    # Save the dictionary to a JSON file
    url = f"https://api.dryg.net/dagar/v2.1/{year}"
    response = requests.get(url)
    data = response.json()

    with open(f"calendar/{year}.json", "w", encoding="utf-8") as json_file:
        json.dump(data, json_file, ensure_ascii=False)

def load_calendar(year):
    file_path = f"calendar/{year}.json"
    try:
        with open(file_path, "r", encoding="utf-8") as json_file:
            return json.load(json_file)      
    except FileNotFoundError:
        # Not cached yet 
        cache_calendar(year)
        return load_calendar(year)

def get_heldagar_exclusive_sundays(year):
    calendar = load_calendar(year)

    if "dagar" in calendar:
        holidays = [(day["datum"], day["helgdag"]) 
                    for day in calendar["dagar"] 
                        if day["röd dag"]=='Ja' 
                            and "helgdag" in day]
        return holidays
    return []

def create_heldagsafton_check(day_name):
    return lambda x: "helgdagsafton" in x  \
                    and day_name.lower() == x["helgdagsafton"].lower() \
                    and x["arbetsfri dag"] == "Nej"

def create_heldag_check(day_name):
    return lambda x: "helgdag" in x  \
                    and day_name.lower()  == x["helgdag"].lower()
from datetime import date, timedelta

# Hitta påskdagen för det givna året
# Påskdagen är den första söndagen efter den första fullmånen efter vårdagjämningen
# Denna beräkning kräver vanligtvis en mer avancerad algoritm, men här används en förenklad version
def caculate_paskdagen(year):
    a = year % 19
    b = year // 100
    c = year % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    L = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * L) // 451
    month = (h + L - 7 * m + 114) // 31
    day = ((h + L - 7 * m + 114) % 31) + 1
    return date(year, month, day)

def calculate_kristi_himmelsfardsdag(year):
    paskdagen = caculate_paskdagen(year)
    # Lägg till 39 dagar för att beräkna Kristi himmelsfärdsdag
    # Kristi himmelsfärdsdag infaller på den 40:e efter påskdagen
    himmelsfardsdag = paskdagen + timedelta(days=39)
    return himmelsfardsdag

def calculate_trettondagsafton(year):
    return date(year, 1, 5)  # Trettondagsafton är alltid den 5 januari

def calculate_skartorsdagen(year):
    paskdagen = caculate_paskdagen(year)
    # Skärtorsdagen är två dagar före påskdagen
    skartorsdagen = paskdagen - timedelta(days=2)
    return skartorsdagen

def get_kortdagar(year):
    calendar = load_calendar(year)
    # Följande dagar är kortdagar om de infaller på en arbetsdag
    # trettondagsafton
    # skärtorsdagen
    # valborgsmässoafton
    # dagen före Kristi himmelsfärdsdag
    # dagen fore Alla helgons dag
    
    trettondagsafton = create_heldagsafton_check("trettondagsafton")
    skärtorsdagen = create_heldagsafton_check("skärtorsdagen")
    valborgsmässoafton = create_heldagsafton_check("valborgsmässoafton")
    allhelgonaafton = create_heldagsafton_check("Allhelgonaafton")
    #kristi_himmelsfärdsdag = create_heldag_check("Kristi himmelsfärdsdag")
    
    if "dagar" in calendar:
        shortdays = [(day["datum"], day["helgdagsafton"])  for day in calendar["dagar"] 
                     if trettondagsafton(day)
                        or skärtorsdagen(day)
                        or valborgsmässoafton(day)
                        or allhelgonaafton(day)]
        kristi_himmelsfardsdag = calculate_kristi_himmelsfardsdag(year)
        dag_fore_kristi_himmelsfardsdag = kristi_himmelsfardsdag - timedelta(days=1)
        shortdays.append((dag_fore_kristi_himmelsfardsdag.strftime("%Y-%m-%d"), "dagen före Kristi himmelsfärdsdag"))
        return shortdays
    return []

def get_kortagar_as_dates(year):
    return [(date_utils.parse_date(day_str), day_name) 
                for day_str, day_name 
                    in get_kortdagar(year)]

def get_heldagar_as_dates(year):
    return [(date_utils.parse_date(holiday_str), holiday_description) 
                for holiday_str, holiday_description 
                    in get_heldagar_exclusive_sundays(year)]

def get_heldag_name(date_obj, heldagar_with_names):
    for day, _ in heldagar_with_names:
        if day == date_obj:
            return _
    return None

def get_kortdag_name(date_obj, kortdagar_with_names):
    for day, _ in kortdagar_with_names:
        if day == date_obj:
            return _
    return None

def list_all_images(dir, file_name=None):
    import os
    files = os.listdir(dir)
    file_name_png = file_name + ".png" if file_name else None
    png_files = [file for file in files if file.endswith(".png")]
    png_files = [file for file in png_files if ((file.lower() == file_name_png.lower()) 
                      or (not file_name_png))]
    
    png_files.sort()
    return png_files

HELDAGAR_AND_KORTDAGAR = ['Nyårsdagen', 'Trettondagsafton', 'Trettondedag jul', 'Skärtorsdagen',
  'Långfredagen', 'Påskdagen', 'Annandag påsk', 'Valborgsmässoafton',
    'Första Maj', 'dagen före Kristi himmelsfärdsdag', 'Kristi himmelsfärdsdag',
      'Pingstdagen', 'Sveriges nationaldag', 'Midsommardagen', 'Allhelgonaafton',
        'Alla helgons dag', 'Juldagen', 'Annandag jul']

def get_heldag_image(heldag_name):
    file_name = list_all_images("images", heldag_name)
    if file_name:
        return  Image(f"images/{file_name}")

def get_image_path(day_description):
    return f"{HELDAGAR_AND_KORTDAGAR.index(day_description)}.png"

if __name__ == "__main__":
    for year in range(2029, 2034):
        cache_calendar(year)
    heldagar = get_heldagar_as_dates(2024)
    kortdagar = get_kortagar_as_dates(2024)
    heldagar_kortdagar = heldagar + kortdagar
    heldagar_kortdagar.sort(key=lambda x: x[0])

    for day, descr in heldagar_kortdagar:
        image_name = f"{HELDAGAR_AND_KORTDAGAR.index(descr)}.png"
        print(f"{day} - {descr} {image_name}")