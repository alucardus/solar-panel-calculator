import openpyxl
import json
import requests
from user_data import User
from create_report import create_report


def main():
    file = "Data.xlsx"
    user = get_input_data(file)
    report = calculator(user)
    create_report(user, report)

# Load input data file, create User object and save information from Excel file to User object
def get_input_data(f):
    wb = openpyxl.load_workbook(f)
    ws = wb.active
    user_data = []
    for row in ws["B3:B11"]:
        for cell in row:
            user_data.append(cell.value)
    user = User(*user_data)
    return user

# Request data using The National Renewable Energy Laboratory (NREL) api
def request_output(u, sq):
    api_key = "K3M1o6DCg2T7VgNB9hEouRMO6rzDgwMguDEencar"
    system_capacity = sq * 0.15
    url = f"https://developer.nrel.gov/api/pvwatts/v8.json?api_key={api_key}&azimuth={u.azimuth}&system_capacity={system_capacity}&losses=14&array_type={u.get_array_key()}&module_type={u.get_module_key()}&gcr=0.4&dc_ac_ratio=1.2&inv_eff=96.0&radius=0&dataset=nsrdb&tilt=20&address={u.address} {u.city} {u.state} {u.country}&soiling=&albedo=0.3&bifaciality=0.7"
    response = requests.get(url)
    data = response.json()
    u.annual_output = data["outputs"]["ac_annual"]
    u.monthly_output = data["outputs"]["ac_monthly"]

# Request electricity price using The National Renewable Energy Laboratory (NREL) api
def request_price(u):
    api_key = "K3M1o6DCg2T7VgNB9hEouRMO6rzDgwMguDEencar"
    url = f"https://developer.nrel.gov/api/utility_rates/v3.json?api_key={api_key}&address={u.address} {u.city} {u.state} {u.country}"
    response = requests.get(url)
    data = response.json()
    u.price = data["outputs"]["residential"]

# Calculate all data for Solar Panel Report
def calculator(u):
    annual_yield = 955  # Annual yield of 1kW solar panel
    total_solar_power = (
        u.annual_consumption / annual_yield
    )  # Total power of solar panels
    panel_size = 1.925  # Size of one solar panel
    square = panel_size / 400 * 1000  # Square of 1kW solar panels
    total_square = total_solar_power * square  # Total square of solar panels
    total_panel_number = total_square / panel_size  # Total number of solar panels
    standard_cost = 750  # Solar panel cost per 1kW
    premium_cost = 1500
    thin_cost = 1000
    installation_cost = 300  # Installation cost of 1 solar panel
    # Total solar panel cost including installation
    if u.module_type == "Standard":
        total_cost = (
            total_solar_power * standard_cost + installation_cost * total_panel_number
        )
    elif u.module_type == "Premium":
        total_cost = (
            total_solar_power * premium_cost + installation_cost * total_panel_number
        )
    else:
        total_cost = (
            total_solar_power * thin_cost + installation_cost * total_panel_number
        )
    # Total solar energy output cost
    request_output(u, total_square)
    request_price(u)
    total_output_cost = u.annual_output * u.price
    # Total repayment period of investments
    repayment_period = total_cost / total_output_cost
    report = [
        total_solar_power,
        total_square,
        total_panel_number,
        total_cost,
        total_output_cost,
        repayment_period,
    ]
    return report


if __name__ == "__main__":
    main()
