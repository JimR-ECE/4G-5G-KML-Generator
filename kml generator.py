import pandas as pd
import simplekml
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime, timedelta
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.styles.stylesheet")


# Function definitions from your second script
def create_sector(longitude, latitude, azimuth, radius):
    """
    Create coordinates for a sector in the form of a pie wedge based on the given parameters.
    """
    if radius is None:
        return None  # Skip this sector as it has no radius

    coords = []
    radius = float(radius)

    # Function to find a point given a start, a bearing, and a distance
    from math import cos, sin, radians, degrees, asin, sqrt, atan2

    def destination_point(lon, lat, bearing, distance):
        R = 6371e3  # Earth's radius in meters
        δ = distance / R  # Angular distance in radians

        θ = radians(bearing)

        φ1 = radians(lat)
        λ1 = radians(lon)

        sin_φ1 = sin(φ1)
        cos_φ1 = cos(φ1)
        sin_δ = sin(δ)
        cos_δ = cos(δ)
        sin_θ = sin(θ)
        cos_θ = cos(θ)

        sin_φ2 = sin_φ1 * cos_δ + cos_φ1 * sin_δ * cos_θ
        φ2 = asin(sin_φ2)
        y = sin_θ * sin_δ * cos_φ1
        x = cos_δ - sin_φ1 * sin_φ2
        λ2 = λ1 + atan2(y, x)

        return degrees(λ2), degrees(φ2)

    # Add the center point
    coords.append((longitude, latitude))

    # Calculate the two edge bearings of the sector based on the azimuth and beamwidth
    start_bearing = (azimuth - 30 / 2) % 360
    end_bearing = (azimuth + 30 / 2) % 360

    # Create the wedge points by finding the destination from the center point using the bearings
    if start_bearing <= end_bearing:
        bearings = range(int(start_bearing), int(end_bearing) + 1)
    else:
        bearings = list(range(int(start_bearing), 360)) + list(range(0, int(end_bearing) + 1))

    for bearing in bearings:
        edge_lon, edge_lat = destination_point(longitude, latitude, bearing, radius)
        coords.append((edge_lon, edge_lat))

    # Close the wedge by connecting back to the center
    coords.append((longitude, latitude))

    return coords


def set_sector_style(pol, color):
    """
    Set the style of the sector based on the color provided in the dataframe.
    Assumes that the color column contains valid KML color strings.
    """
    pol.style.linestyle.color = color  # Set the border color
    pol.style.linestyle.width = 2
    pol.style.polystyle.color = simplekml.Color.changealphaint(0, simplekml.Color.white)  # No fill


def create_description(row):
    """
    Create a description string from the row data.
    """
    description = f"Cell Name: {row['cellname']}<br>"
    description += f"Frequency: {row['Freq']}<br>"
    description += f"Site_ID: {row['SiteID']}<br>"
    description += f"Cell ID: {row['CellID']}<br>"
    description += f"PCI: {row['PCI']}<br>"
    description += f"EARFCN: {row['EARFCN']}<br>"
    description += f"Height: {row['HT']}<br>"
    description += f"Azimuth: {row['Azimuth']}<br>"
    description += f"Mechanical Tilt: {row['MTILT']}<br>"
    description += f"Remote Electrical Tilt: {row['ETILT']}<br>"
    description += f"OAM IP: {row['OAM IP']}<br>"
    return description


# Mapping of common color names to KML hexadecimal color strings
color_mapping = {
    'blue': 'ffff0000',  # Opaque blue
    'red': 'ff0000ff',  # Opaque red#
    'lime': 'ff00ff00',
    'yellow': 'ff00ffff',
    'orange': 'ff00a5ff',
    'cyan': 'ffffff00',
    'purple': 'ff800080'
    # Add more mappings as needed
}


def get_kml_color(color_name):
    """
    Convert a color name to a KML hexadecimal color string.
    If the color name is not found in the mapping, default to white ('ffffffff').
    """
    return color_mapping.get(color_name.lower(), 'ffffffff')


def create_kml(dataframe, filename):
    kml = simplekml.Kml()
    created_sites = set()  # Initialize the set here
    freq_folders = {}  # Dictionary to hold the folders for each frequency
    ibs_folder = kml.newfolder(name='IBS')  # Create a folder for IBS sites

    for index, row in dataframe.iterrows():
        site_type = row['Site Type']

        # Handle 'IBS' sites by creating a folder and adding points without icons
        if site_type == 'IBS':
            if row['SiteName'] not in freq_folders:  # Check if the site name is already added
                pnt = ibs_folder.newpoint(name=row['SiteName'], coords=[(row['longitude'], row['latitude'])])
                pnt.style.iconstyle.icon.href = ''  # Disable the icon for IBS sites
                pnt.style.iconstyle.scale = 0  # Optionally set the scale to 0
                pnt.style.labelstyle.scale = 0.8  # Adjust label scale if needed
                freq_folders[row['SiteName']] = pnt  # Remember that this site is already added
            continue  # Skip the rest of the loop for 'IBS' sites

        freq = row['Freq']
        # Handle other sites as before
        if freq not in freq_folders:
            freq_folders[freq] = kml.newfolder(name=f"{freq}")

        # Create the sector polygon within the appropriate frequency folder
        coords = create_sector(row['longitude'], row['latitude'], row['Azimuth'], row['Radius'])
        if coords is None:
            continue  # Skip this row as it has no valid sector

        pol = freq_folders[freq].newpolygon(outerboundaryis=coords)
        pol.description = create_description(row)

        # Set the polygon style using the color from the 'Color' column
        kml_color = get_kml_color(row['Color'])
        pol.style.linestyle.color = kml_color
        pol.style.linestyle.width = 2
        pol.style.polystyle.color = simplekml.Color.changealphaint(0, simplekml.Color.white)  # No fill

        # Add a label for the site if it hasn't been added yet
        site_name = row['SiteName']
        #site_name = row['SiteName'] + ' \n ' + str(row['SiteID'])

        if site_name not in created_sites:
            pnt = freq_folders[freq].newpoint(name=site_name, coords=[(row['longitude'], row['latitude'])])
            pnt.style.iconstyle.icon.href = ''  # This disables the icon
            pnt.style.iconstyle.scale = 0  # Optionally set the scale to 0
            pnt.style.labelstyle.color = simplekml.Color.white  # Set label color to white
            pnt.style.labelstyle.scale = 0.8  # Set label scale to 0.8 or as desired
            created_sites.add(site_name)

    kml.save(filename)


# Function to process and combine the 4G and 5G databases
def process_databases():
    root = tk.Tk()
    root.withdraw()

    # Display a trademark message during file import
    messagebox.showinfo("KML Generator Tool - Created by Jim Ramos", "Input the 4G and 5G Engineering Database")

    # File selection dialogs for 4G and 5G Excel files
    file_4g_path = filedialog.askopenfilename(title='Select the 4G Excel file')
    if not file_4g_path:  # Check if a file was selected
        messagebox.showerror("Error", "No 4G file selected. Exiting.")
        return None

    file_5g_path = filedialog.askopenfilename(title='Select the 5G Excel file')
    if not file_5g_path:  # Check if a file was selected
        messagebox.showerror("Error", "No 5G file selected. Exiting.")
        return None

    # Read the data from both files into pandas DataFrames
    df_4g = pd.read_excel(file_4g_path)
    df_5g = pd.read_excel(file_5g_path)

    # Define a mapping of the old column names to the new column names for 4G and 5G
    column_mappings = {
        '4G': {
            'Physical Site ID': 'SiteName',
            'EnodeB ID': 'SiteID',
            'Cell ID': 'CellID',
            'Cell Name': 'cellname',
            'Cell Longitude': 'longitude',
            'Cell Latitude': 'latitude',
            'PCI(Physical Cell Identifier)': 'PCI',
            'Downlink Frequency': 'EARFCN',
            'Frequency Band': 'Freq',
            'Antenna Height': 'HT',
            'Azimuth Angle': 'Azimuth',
            'Mechanical Downtilt': 'MTILT',
            'Electrical Downtilt': 'ETILT',
            'Site Type': 'Site Type',
            'Oam IP': 'OAM IP',
        },
        '5G': {
            'Physical Site ID': 'SiteName',
            'GnodeB ID': 'SiteID',
            'Cell ID': 'CellID',
            'Cell Name': 'cellname',
            'Cell Longitude': 'longitude',
            'Cell Latitude': 'latitude',
            'Physical Cell ID': 'PCI',
            'Downlink Frequency': 'EARFCN',
            'Frequency Band': 'Freq',
            'Antenna Height': 'HT',
            'Azimuth Angle': 'Azimuth',
            'Mechanical Downtilt': 'MTILT',
            'Electrical Downtilt': 'ETILT',
            'Site Type': 'Site Type',
            'OAM IP': 'OAM IP',
        }
    }

    # Select and rename the columns for the 4G DataFrame
    df_4g_selected = df_4g[list(column_mappings['4G'].keys())]
    df_4g_renamed = df_4g_selected.rename(columns=column_mappings['4G'])

    # Select and rename the columns for the 5G DataFrame
    df_5g_selected = df_5g[list(column_mappings['5G'].keys())]
    df_5g_renamed = df_5g_selected.rename(columns=column_mappings['5G'])

    # Combine the two DataFrames into one
    combined_df = pd.concat([df_4g_renamed, df_5g_renamed], ignore_index=True)

    # Function to determine radius and color based on EARFCN
    def determine_radius_color(earfcn):
        radius_color_map = {
            9610: ('250', 'blue'),
            425: ('200', 'red'),
            450: ('200', 'red'),
            2155: ('200', 'red'),
            550: ('175', 'purple'),
            40140: ('150', 'lime'),
            623334: ('100', 'yellow'),
            633334: ('80', 'orange'),
        }
        return radius_color_map.get(earfcn, (None, None))  # Return None if EARFCN is not in the map

    # Apply the function to the 'EARFCN' column to create new 'Radius' and 'Color' columns
    combined_df[['Radius', 'Color']] = combined_df.apply(lambda row: determine_radius_color(row['EARFCN']),
                                                         axis=1,
                                                         result_type='expand')

    # Filter the DataFrame for 'Macro', 'Micro', and 'IBS' sites
    filtered_df = combined_df[combined_df['Site Type'].isin(['Macro', 'Micro', 'IBS'])]

    # Modified function to determine radius and color, including logic for Micro sites
    def determine_radius_color_adjusted(row):
        if row['Site Type'] == 'Micro':
            return ('50', 'cyan')  # Fixed radius and color for Micro sites
        else:
            return determine_radius_color(row['EARFCN'])  # Existing logic for Macro sites

    # Apply the modified function using .loc
    radius_color = filtered_df.apply(determine_radius_color_adjusted, axis=1, result_type='expand')
    filtered_df.loc[:, 'Radius'] = radius_color[0]
    filtered_df.loc[:, 'Color'] = radius_color[1]

    return filtered_df


def check_expiration():
    # Hardcoded start date in format YYYY-MM-DD
    start_date_str = '2024-03-15'  # Replace with your desired start date
    expiration_period = timedelta(days=60)  # 3 months

    # Convert string to date
    start_date = datetime.strptime(start_date_str, '%Y-%m-%d')

    # Check if the current date is beyond the expiration period
    if datetime.now() > start_date + expiration_period:
        messagebox.showerror("Expired", "This application has expired. Please contact Jim Ramos")
        return False
    return True


# Main function to run the process
def run_process():
    combined_df = process_databases()
    if combined_df is None:
        return  # Exit the function if no data was processed

    output_kml_path = filedialog.asksaveasfilename(defaultextension=".kml",
                                                   filetypes=[("KML files", "*.kml")],
                                                   title="Save the KML file")
    if output_kml_path:
        create_kml(combined_df, output_kml_path)
        messagebox.showinfo("Success", f"KML file created successfully at: {output_kml_path}")


# Run the process
if __name__ == "__main__":
    if check_expiration():
        run_process()

