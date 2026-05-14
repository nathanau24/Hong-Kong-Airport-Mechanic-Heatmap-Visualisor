Hong Kong International Airport Mechanic Heatmap Visualiser

This project provides a MATLAB-based tool for visualising and analysing mechanic workload and distribution across an airport. By processing flight arrival and departure schedules, the system generates an animated heatmap that highlights peak demand periods and resource allocation requirements.

Overview:
The core of the project is the `generate_mechanic_heatmap_animation.m` script. It simulates a 24-hour period divided into 5-minute intervals (288 intervals total) to estimate the number of mechanics required at various parking bays based on airline-specific rules and flight activity.

Key Features
- Dynamic Heatmap Animation: Visualises the flow of ground personnel across the airport layout over time.
- Late/Early Logic: Accounts for schedule variations in arrivals and departures to provide more accurate resource estimates.
- Airline Grouping Rules: Applies specific mechanic allocation counts based on the operating airline (e.g., CI, MU, CA, etc.).
- eak Statistics Analysis: Automatically identifies the busiest 5-minute intervals and calculates the daily average peak mechanic requirements airport-wide.

File Structure
- `generate_mechanic_heatmap_animation.m`: The primary MATLAB script for data processing and visualisation.
- `Animation_Data.xlsx`: Contains the 'Arrival' and 'Departure' flight schedules used for the simulation.
- `Aerodrome Map.xlsx`: Defines the physical layout of the airport bays and provides the templates for heatmap generation.

Usage
1.  Ensure all project files are located in the same directory within MATLAB.
2.  Open `generate_mechanic_heatmap_animation.m`.
3.  Run the script.
4.  The console will output the "Busiest 5-Minute Intervals" and the "Average Peak Mechanics" required, while a figure window will display the animated heatmap.

Requirements
- MATLAB (with Image Processing or Statistics toolboxes recommended for full functionality).
- Microsoft Excel (to manage the input data files).

---
Developed by Nathan Au, August 2025.
