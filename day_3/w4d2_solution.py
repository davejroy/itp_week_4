# need my imports

import requests
import openpyxl
import json

pokemon_list = []

# function that returns a list of pokemon name and api to call for skills
# THIS FUNCTION TAKES A LONG TIME****
def populate_pokemon(url):
    response = requests.get(url)
    json_data = json.loads(response.text)
    for each in json_data['results']:
        pokemon_list.append(each)
    # BONUS: THIS IS AN EXIT CASE, NEEDED FOR RECURSIVE FUNCTIONS
    if json_data['next'] != None:
        # BONUS: THIS IS A RECURSIVE FUNCTION
        populate_pokemon(json_data['next'])

# function that returns the string of the abilities
# from a list of abilities provided from the API
def stringify_abilities(abilities_list):
    result = ""
    first = True
    for ability in abilities_list:
        if first is True:
            result += ability['ability']['name']
            first = False
        else:
            result += ", " + ability['ability']['name']
    return result

def retrieve_abilities(url):
    response = requests.get(url)
    json_data = json.loads(response.text)
    return json_data['abilities']

# function that takes every pokemon retrieve the skills list
def process_pokemon_data():
    for idx, pokemon in enumerate(pokemon_list, start=1):
        name = pokemon['name']
        abilities_list = retrieve_abilities(pokemon['url'])
        abilities_string = stringify_abilities(abilities_list)
        add_row(idx, name, abilities_string)

# define my workbook
wb = openpyxl.Workbook()
sheet = wb.active

# function that takes in row_num, name, and skills
# skills need to be a string of all the skills comma separated
# returns nothing
def add_row(row_num, name, skills):
    sheet['A' + str(row_num)] = name
    sheet['B' + str(row_num)] = skills

# final sequence of commands
populate_pokemon("https://pokeapi.co/api/v2/pokemon/")
process_pokemon_data()
wb.save("/home/dkayzee/vit/intro-python-august-2021/itp_week_4/day_3/output.xlsx")