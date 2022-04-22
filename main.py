
import openpyxl
import numbers
import openpyxl.utils
import plotly.graph_objects
from state_abbrev import us_state_to_abbrev

# opens up the excel file and returns the data
def open_worksheet(filename):
    county_pop = openpyxl.load_workbook(filename)
    data_sheet = county_pop.active
    return data_sheet

def main():
    pop_worksheet = open_worksheet('countyPopChange2020-2021.xlsx')
    show_pop_change = should_display_pop_change()
 #depending on answer, either show_pop_change_map will appear or show_percent_change_map
    if show_pop_change is True:
        return show_pop_change_map(pop_worksheet)
    else:
        return show_percent_change_map(pop_worksheet)

#asks the question of if the map should be dislayed
def should_display_pop_change():
    user_response = input("Should I display a map of total population changes? ")
    response = user_response.lower()
    good_answers = ["yes", "sure", "fine", "ok"]
    if response not in good_answers:
        return False
    else:
        return True

def show_pop_change_map(pop_sheet):
    list_of_state_abbrev = []
    list_of_pop_changes = []
    for row in pop_sheet.rows:
        county_cell = row[4]
        pop_cell = row[11]
        pop_value = pop_cell.value
        county_value = county_cell.value
        if isinstance(pop_value, numbers.Number):
            continue
        if county_cell.value not in us_state_to_abbrev:
            continue
        if county_value == 0:
            state_name = county_cell.value
            state_abbrev = us_state_to_abbrev[state_name]
            list_of_state_abbrev.append(state_abbrev)
            pop_estimate2021_cell_number = openpyxl.utils.cell.column_index_from_string('n') - 1
            pop_estimate2021_cell = row[pop_estimate2021_cell_number]
            pop_estimate2021 = pop_estimate2021_cell.value
            pop_value = pop_value - pop_estimate2021
            list_of_pop_changes.append(pop_value)
#brings up the map of the population change

        map_to_show = plotly.graph_objects.Figure(
            data=plotly.graph_objects.Choropleth(
                locations=list_of_state_abbrev,
                z=list_of_pop_changes,
                locationmode="USA-states",
                colorscale='picnic',
                colorbar_title="population change"
            )
        )
        map_to_show.update_layout(
            title_text="population change",
            geo_scope="usa"

        )
        map_to_show.show()

def show_percent_change_map(pop_sheet):
    list_of_state_abbrev = []
    list_of_pop_changes = []
    for row in pop_sheet.rows:
        fourth_cell = row[4]
        pop_cell = row[9]
        pop2020 = row[11]
        pop_estimate = row[13]
        pop_value = pop_cell.value
        fourth_value = fourth_cell.value
        pop2020_value = pop2020.value
        if isinstance(pop_value, numbers.Number):
            continue
        if fourth_cell.value not in us_state_to_abbrev:
            continue
        if fourth_value == 0:
            state_name = fourth_cell.value
            state_abbrev = us_state_to_abbrev[state_name]
            list_of_state_abbrev.append(state_abbrev)
            pop_estimate2021_cell_number = openpyxl.utils.cell.column_index_from_string('n') - 1
            pop_estimate2021_cell = row[pop_estimate2021_cell_number]
            pop_estimate2021 = pop_estimate2021_cell.value
            pop_change = pop_value - pop_estimate2021
            list_of_pop_changes.append(pop_change)
#brings up the map of the percent change

    map_to_show = plotly.graph_objects.Figure(
        data=plotly.graph_objects.Choropleth(
            locations=list_of_state_abbrev,
            z=list_of_pop_changes,
            locationmode="USA-states",
            colorscale='picnic',
            colorbar_title="population change"
        )
    )
    map_to_show.update_layout(
        title_text="population change",
        geo_scope="usa"

    )
    map_to_show.show()

main()