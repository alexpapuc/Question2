from pathlib import Path
from win32com.client import Dispatch
import re

#set the path where the information is stored
file_path = Path.cwd() / "question2.xlsx"


#run Excel App, read worksheet no.1 from "question2.xlsx"
xl = Dispatch('Excel.Application')
wb = xl.Workbooks.Open(Filename = file_path)
ws = wb.Worksheets(1)


# save the string from excel text box into a variable
text_string = ws.Shapes(1).TextFrame.Characters().Text


#check that out data is a string
#print(type(text_string))


#close Excel App
xl.Quit()

def use_regex(text_string):
    """this function takes text_string as an argument and compile the pattern created in order to return the
    match based on patern criteria.
    Pattern explanation
    We are looking to create a pattern that match the following string:
    kappa 6.2.0 (Release) Mar 15 2021
    kappa ->[a-zA-Z]+\s -> Recognizes multiple characters and one whitespace
    6.2.0 ->([0-9]+(\.[0-9]+)+)\s -> Recognizes combination [number+character] and one whitespace
    (Release) ->\([a-zA-Z]+\)\s -> Recognizes pharantheses, multiple characters,pharantheses and one whitespace
    Mar ->[a-zA-Z]+\s -> Recognizes multiple characters and one whitespace
    15 ->[0-9]+\s -> Recognizes number and one whitespace
    2021->[0-9]+ -> Recognizes number """
    pattern = re.compile(r"[a-zA-Z]+\s([0-9]+(\.[0-9]+)+)\s\([a-zA-Z]+\)\s[a-zA-Z]+\s[0-9]+\s[0-9]+")
    return pattern.search(text_string)


os_version = print(f'OS Version is: {use_regex(text_string)[0]}')

if __name__ == "__main__":
    use_regex(text_string)

