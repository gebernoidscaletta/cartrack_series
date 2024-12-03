import json
import openpyxl

# File paths
input_file = ""  # Replace with your JSON file path
output_file = ""

# Load JSON data from file
with open(input_file, "r") as file:
    data = json.load(file)

# Initialize an Excel workbook
wb = openpyxl.Workbook()
sheet = wb.active
sheet.title = "Geofence Data"

# Write headers
headers = ["Geofence Name", "Geofence Description", "Points", "Group Name", "Group Description"]
sheet.append(headers)

# Process JSON data
for feature in data["features"]:
    no_petak = feature["properties"]["NO_PETAK"]
    description = f"Nomor Petak - {no_petak}"

    # Check geometry type and extract coordinates
    geometry_type = feature["geometry"]["type"]
    points = []

    if geometry_type == "Polygon":
        for coord in feature["geometry"]["coordinates"][0]:
            lat, lon = coord[1], coord[0]  # Latitude, Longitude
            points.append(f"{lat}#{lon}")

    elif geometry_type == "MultiPolygon":
        for polygon in feature["geometry"]["coordinates"]:
            for coord in polygon[0]:
                lat, lon = coord[1], coord[0]  # Latitude, Longitude
                points.append(f"{lat}#{lon}")

    # Convert points to the required format
    points_str = "#".join(points)

    # Group data
    group_name = "Batas Petak KTH"
    group_description = "None"

    # Append data to Excel sheet
    sheet.append([no_petak, description, points_str, group_name, group_description])

# Save the Excel file
wb.save(output_file)
print(f"Data successfully written to {output_file}")
