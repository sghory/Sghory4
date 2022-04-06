import openpyxl
import numbers
import openpyxl.utils

# cannot get the information from the sheets

def open_worksheet(filename):
    county_pop = openpyxl.load_workbook(filename)
    data_sheet = county_pop.active
    return data_sheet
    # opens Excel file

def main():
    county_pop_sheet = open_worksheet("countyPopChange2020-2021.xlsx")
    county_pop_losses = should_get_losses(county_pop_sheet)
    # gets information from Excel sheet

def should_get_losses(process_data):
    # answer = int(input("Should I get the counties that lost population?"))
    # print(f"Counties that lost population {answer} in 2021")
    # asks the question
    response = ""
    response = input("Should I get the counties that lost population in 2021? ")
    response = response.lower()
    good_answers = ["yes", "sure", "fine", "ok"]
    if response not in good_answers:
        return False
    else:
        return True

def process_data(pop_sheet,show_losses):
    list_of_pop_changes = []
    county_pop_losses = should_get_losses(county_pop_sheet)
    for row in pop_sheet.rows:
        county_cell = row[6]
        pop_cell = row[9]
        pop_change = row[10]
        state_name = row[5]
        pop_value = pop_cell.value
        if not isinstance(pop_value, numbers.Number):
            continue
        pop_estimate2021_cell_number = openpyxl.utils.cell.column_index_from_string('n') - 1
        pop_estimate2021_cell = row[pop_estimate2021_cell_number]
        pop_estimate = pop_estimate2021_cell.value
        pop_difference = pop_change / pop_cell
        if show_losses and pop_difference < -2.5:
            print(f"{pop_cell.value} : {pop_value}")
        if show_losses and pop_difference > -1.5:
            print(f"{pop_cell.value} : {pop_value}")
        list_of_pop_changes.append(pop_difference)
    # change/pop-change (x 100) for percent change
    # calculates the population change

main()