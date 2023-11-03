
# SOLAR PANEL CALCULATOR

This my Solar Panel Calculator - final project for CS50P course by Harvard university.  This program, based on input data from the user (for example, address, roof area, electricity consumption), calculates required number of solar panels, annual amount of electricity generated, total cost of installation  and payback period of the investment.
#### This version of Solar Panel Calculator works only for US citizens.

## Prerequisites

Before you begin, ensure you have met the following requirements:
* You have installed the latest version of Python
* You have installed openpyxl v3.1.2
* You have installed pytest v7.4.2
* You have installed requests v2.31.0

## Using Solar Panel Calculator

To use Solar Panel Calculator, follow these steps:

* download all files, including Excel file
* open `Data.xlsx` file and fill in all required information ([How to fill in Data.xlsx file](##how-to-fill-in-data.xlsx-file)), save this file and close it
* run `project.py`
* new Excel file will be created in your folder called `Report.xlsx`. This file will contain all output data for current user (e.g. required number of solar panels, annual amount of electricity generated, total cost of installation, payback period of the investment)

#### Video Demo:  <https://www.youtube.com/watch?v=3LLwVofu_vQ&ab_channel=AndrejsSkripko>

## Description
My project consists of the following files:
- `create_data_file.py`: creates Excel file `Data.xlsx`, where user can fill in required information about his household
- `user_data.py`: creates class User and saves information about current user
- `project.py`: this is my main programm which loads data from `Data.xlsx` file, requests all necessary data using The National Renewable Energy Laboratory (NREL) api and then calculate all Solar Panel Report data
- `create_report.py`: creates Excel file `Report.xlsx` and fill in all data from `project.py` file
- `test_project.py`: tests all my files using `pytest` library

## How to fill in Data.xlsx file
`Data.xlsx` file contains 9 input fields. Some fields are optional, some fields are required:
* Address (optional, full address of your household)
* City (required, e.g. Austin)
* State (required, e.g. TX or Texas)
* Country (optional, e.g. USA or United States)
* Rooftop size, m2 (required, e.g. 200)
* Annual electricity consumption, kWh (required, e.g. 11000)
* Solar panel module type (choose from dropdown list: Standard, Premium, Thin film). See detailed information [here](####solar-panel-module-type)
* Fixing array type (choose from dropdown list: Fixed (open track), Fixed (roof mount), 1-Axis, 1-Axis Backtracking, 2-Axis). See detailed information [here](####fixing-array-type)
* Azimuth of rooftop (choose from dropdown list: 0, 45, 90, 135, 180, 225, 270, 315). See detailed information [here](####azimuth)


#### Solar panel module type
The module type describes the photovoltaic modules in the array. If you do not have information about the modules in the system, use the default Standard module type. Otherwise, you can use information from the module data sheet and the table below to choose the module type.


| PVWatts® Module Type | Cell Material       | Approximate Nominal Efficiency | Module Cover                       | Temperature Coefficient of Power |
| -------------------- | ------------------- | ------------------------------ | ---------------------------------- | -------------------------------- |
| Standard             | Crystalline Silicon | 19%                            | Glass with anti-reflective coating | \-0.37%/°C                       |
| Premium              | Crystalline Silicon | 21%                            | Glass with anti-reflective coating | \-0.35%/°C                       |
| Thin Film            | Thin film           | 18%                            | Glass with anti-reflective coating | \-0.32%/°C                       |

#### Fixing array type
The array type describes whether the PV modules in the array are fixed, or whether they move to track the movement of the sun across the sky with one or two axes of rotation. The default value is for a fixed array, or one with no tracking. Use the following diagrams to choose the appropriate option for your system:

![alt text](https://pvwatts.nrel.gov/docs/en/tracking.png)

Fixed Open and Roof Mount Options
For systems with fixed arrays, you can choose between an open rack or a roof mount option. The open rack option is appropriate for ground-mounted systems. It assumes that air flows freely around the array, helping to cool the modules and reduce cell operating temperatures. (The array's output increases as the cell temperature decreases at a given incident solar irradiance.) The roof mount option is typical of residential installations where modules are attached to the roof surface with standoffs that provide limited air flow between the module back and roof surface.

For the open rack option, assumed an installed nominal operating temperature (INOCT) of 45 degrees Celsius. For roof mount systems, the INOCT is 50 degrees Celsius, which corresponds roughly to a three or four inch standoff height.

One-axis Tracking and Backtracking

For the one-axis tracking option, programm models the effect of self-shading. Self-shading is a reduction in the array's output caused by shading of neighboring rows of modules at certain times of day and year when the sun is low on the horizon. One-axis tracking with backtracking is a tracking algorithm that rotates the array toward the horizontal during early morning and late evening hours to reduce the effect of self shading.

The self-shading calculations use the advanced parameter Ground Coverage Ratio (GCR) to represent the distance between rows of modules in the array, and assume a rotation limit of ±45 degrees from the horizontal.

#### Azimuth
For a fixed array, the azimuth angle is the angle clockwise from true north describing the direction that the array faces. An azimuth angle of 180° is for a south-facing array, and an azimuth angle of zero degrees is for a north-facing array.

For an array with one-axis tracking, the azimuth angle is the angle clockwise from true north of the axis of rotation. The azimuth angle does not apply to arrays with two-axis tracking.

The default value is an azimuth angle of 180° (south-facing) for locations in the northern hemisphere and 0° (north-facing) for locations in the southern hemisphere. These values typically maximize electricity production over the year, although local weather patterns may cause the optimal azimuth angle to be slightly more or less than the default values. For the northern hemisphere, increasing the azimuth angle favors afternoon energy production, and decreasing the azimuth angle favors morning energy production. The opposite is true for the southern hemisphere.
| Heading |  Azimuth  |
|:-----|:--------:|
| N   | 0 |
| NE   |  45  |
| E   | 90 |
| SE   | 135 |
| S   | 180 |
| SW   | 225 |
| X   | 270 |
| NW   | 315 |

## Contact

* If you want to contact me you can reach me at <andrejs.skripko@gmail.com>.
* Check my YouTube channel: https://www.youtube.com/channel/UCwWH0b6FUnoGb2AAyhnhq5w
