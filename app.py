import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom
from datetime import datetime
import os

st.title("Excel to FPL Converter")

# File uploader
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Get Excel filename without extension
    base_filename = os.path.splitext(uploaded_file.name)[0]
    
    # Read the Excel file and get the sheet name
    excel_data = pd.ExcelFile(uploaded_file)
    sheet_name = excel_data.sheet_names[0]
    
    # Read the first sheet
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
    
    # Check required columns
    required_cols = ['Longitude', 'Latitude', 'Description']
    if not all(col in df.columns for col in required_cols):
        st.error(f"Excel must contain {', '.join(required_cols)} columns!")
    else:
        # Clean data
        df = df.dropna(subset=required_cols)
        df['Longitude'] = df['Longitude'].apply(lambda x: f"{float(x):.4f}")
        df['Latitude'] = df['Latitude'].apply(lambda x: f"{float(x):.4f}")
        df['Description'] = df['Description'].astype(str)

        # Split into chunks of 99 waypoints
        max_waypoints = 99
        for i in range(0, len(df), max_waypoints):
            chunk_df = df.iloc[i:i + max_waypoints]
            
            # Create XML structure
            flight_plan = ET.Element("flight-plan", xmlns="http://www8.garmin.com/xmlschemas/FlightPlan/v1")
            ET.SubElement(flight_plan, "created").text = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
            
            waypoint_table = ET.SubElement(flight_plan, "waypoint-table")
            for index, row in chunk_df.iterrows():
                waypoint = ET.SubElement(waypoint_table, "waypoint")
                ET.SubElement(waypoint, "identifier").text = row['Description']
                ET.SubElement(waypoint, "type").text = "USER WAYPOINT"
                ET.SubElement(waypoint, "lat").text = row['Latitude']
                ET.SubElement(waypoint, "lon").text = row['Longitude']
                ET.SubElement(waypoint, "comment").text = row['Description']

            route = ET.SubElement(flight_plan, "route")
            ET.SubElement(route, "route-name").text = sheet_name
            ET.SubElement(route, "flight-plan-index").text = "1"
            
            for _, row in chunk_df.iterrows():
                route_point = ET.SubElement(route, "route-point")
                ET.SubElement(route_point, "waypoint-identifier").text = row['Description']
                ET.SubElement(route_point, "waypoint-type").text = "USER WAYPOINT"

            # Convert to string with proper XML declaration and pretty print
            xml_string = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            rough_string = ET.tostring(flight_plan, encoding='utf-8', method='xml')
            reparsed = minidom.parseString(rough_string)
            xml_string += reparsed.toprettyxml(indent="  ")[23:]  # Skip XML declaration

            # Provide download link for each chunk
            file_name = f"{base_filename}{'_' + str(i//max_waypoints + 1) if i > 0 else ''}.fpl"
            st.download_button(
                label=f"Download {file_name}",
                data=xml_string,
                file_name=file_name,
                mime="application/xml"
            )
            st.success(f"{file_name} is ready to download and use with your SD card!")
