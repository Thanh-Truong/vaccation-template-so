import re

def parse_input_periods(custom_periods):
    try:
        periods_as_tuples = [tuple([int(j) 
                  for j in (e.replace(' ', '').split('-'))]) 
           for e in map(str.strip, custom_periods.split(","))]
        # making tuples of two (start, end) 
        return [(item[0], item[1] if len(item) > 1 else item[0]) 
                   for item in periods_as_tuples]
    except ValueError:
        print("Ogiltig inmatning. Ange ett heltal.")
        raise


if __name__ == "__main__":
    # Example usage when the script is run
    input_string = "  1 -4, 5  -   8 , 9-12, 1, 5, 9"
    print(parse_input_periods(input_string))
    input_string = "  1,2,3,4,5,6,7,8,9,10,11,12"
    for s,e in parse_input_periods(input_string):
        print(s,e)
        