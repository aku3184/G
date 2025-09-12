import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom
from datetime import datetime
import os
import re
import io

st.title("Excel/KML to Multi-Format Converter")

# File uploader，支持KML
uploaded_file = st.file_uploader("Upload your file", type=["xlsx", "xls", "kml"])

# 输出格式选择
output_format = st.selectbox("Select output format", ["FPL", "KML", "GPX", "Excel"])

# 输出经纬度格式选择（只影响Excel输出）
coord_format = st.selectbox("Select output coordinate format", ["Decimal Degrees", "Degrees Minutes (DM)", "Degrees Minutes Seconds (DMS)"])

def parse_coord(s):
    """解析经纬度：支持Decimal, DM, DMS，只支持英文符号和方向"""
    s = str(s).strip().upper().replace('′', "'").replace('″', '"').replace('＇', "'").replace('＂', '"')  # 标准化特殊符号
    try:
        # 直接Decimal
        return float(s)
    except ValueError:
        pass
    
    # 方向映射：只支持英文
    direction_map = {'N': 1, 'S': -1, 'E': 1, 'W': -1, 'Ｎ': 1, 'Ｓ': -1, 'Ｅ': 1, 'Ｗ': -1}
    
    # DMS: 如 40° 26' 46" N 或 -40° 26' 46" S
    match = re.match(r'(-?\d+)°?\s*(\d+)\'\s*(\d+(?:\.\d+)?)"\s*([NSEW]?)$', s)
    if match:
        d, m, sec, dir_str = match.groups()
        dd = abs(float(d)) + float(m)/60 + float(sec)/3600
        sign = -1 if d.startswith('-') else direction_map.get(dir_str, 1)
        return sign * dd
    
    # DM: 如 40° 26.7667' N 或 -40° 26.7667' S
    match = re.match(r'(-?\d+)°?\s*(\d+(?:\.\d+)?)\'\s*([NSEW]?)$', s)
    if match:
        d, m, dir_str = match.groups()
        dd = abs(float(d)) + float(m)/60
        sign = -1 if d.startswith('-') else direction_map.get(dir_str, 1)
        return sign * dd
    
    raise ValueError(f"Invalid coordinate format: {s}")

def dd_to_dms(dd, is_lat=True):
    """Decimal to DMS string，如 40°26'46\"N"""
    negative = dd < 0
    dd = abs(dd)
    minutes, seconds = divmod(dd * 3600, 60)
    degrees, minutes = divmod(minutes, 60)
    direction = ('S' if negative else 'N') if is_lat else ('W' if negative else 'E')
    return f"{int(degrees)}°{int(minutes)}'{seconds:.0f}\"{direction}"

def dd_to_dm(dd, is_lat=True):
    """Decimal to DM string，如 40°26.7667'N"""
    negative = dd < 0
    dd = abs(dd)
    degrees = int(dd)
    minutes = (dd - degrees) * 60
    direction = ('S' if negative else 'N') if is_lat else ('W' if negative else 'E')
    return f"{degrees}°{minutes:.4f}'{direction}"

if uploaded_file is not None:
    base_filename = os.path.splitext(uploaded_file.name)[0]
    file_ext = os.path.splitext(uploaded_file.name)[1].lower()
    
    if file_ext in ['.xls', '.xlsx']:
        excel_data = pd.ExcelFile(uploaded_file)
        sheet_name = excel_data.sheet_names[0]
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
    elif file_ext == '.kml':
        # 解析KML
        uploaded_file.seek(0)  # 重置文件指针
        tree = ET.parse(uploaded_file)
        root = tree.getroot()
        ns = {'kml': 'http://www.opengis.net/kml/2.2'}
        placemarks = root.findall('.//kml:Placemark', ns)
        data = []
        for pm in placemarks:
            name_elem = pm.find('kml:name', ns)
            desc = name_elem.text if name_elem is not None else ''
            coord_elem = pm.find('.//kml:coordinates', ns)
            if coord_elem is not None:
                coords = coord_elem.text.strip().split(',')
                if len(coords) >= 2:
                    lon, lat = coords[0], coords[1]
                    data.append({'Latitude': lat, 'Longitude': lon, 'Description': desc})
        df = pd.DataFrame(data)
        sheet_name = base_filename  # 用文件名作为route-name
    else:
        st.error("Unsupported file type!")
        st.stop()
    
    # 检查必要列
    required_cols = ['Longitude', 'Latitude', 'Description']
    if not all(col in df.columns for col in required_cols):
        st.error(f"File must contain {', '.join(required_cols)} columns!")
    else:
        # 清理数据
        df = df.dropna(subset=required_cols)
        df['Latitude'] = df['Latitude'].apply(parse_coord)
        df['Longitude'] = df['Longitude'].apply(parse_coord)
        df['Description'] = df['Description'].astype(str)
        
        # 分块（如果需要，假设max_waypoints=99只为FPL）
        max_waypoints = 99 if output_format == 'FPL' else len(df)  # 其他格式不限
        for i in range(0, len(df), max_waypoints):
            chunk_df = df.iloc[i:i + max_waypoints]
            
            if output_format == 'FPL':
                # 原FPL逻辑
                flight_plan = ET.Element("flight-plan", xmlns="http://www8.garmin.com/xmlschemas/FlightPlan/v1")
                ET.SubElement(flight_plan, "created").text = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
                waypoint_table = ET.SubElement(flight_plan, "waypoint-table")
                for _, row in chunk_df.iterrows():
                    waypoint = ET.SubElement(waypoint_table, "waypoint")
                    ET.SubElement(waypoint, "identifier").text = row['Description']
                    ET.SubElement(waypoint, "type").text = "USER WAYPOINT"
                    ET.SubElement(waypoint, "lat").text = f"{row['Latitude']:.4f}"
                    ET.SubElement(waypoint, "lon").text = f"{row['Longitude']:.4f}"
                    ET.SubElement(waypoint, "comment").text = row['Description']
                route = ET.SubElement(flight_plan, "route")
                ET.SubElement(route, "route-name").text = sheet_name
                ET.SubElement(route, "flight-plan-index").text = "1"
                for _, row in chunk_df.iterrows():
                    route_point = ET.SubElement(route, "route-point")
                    ET.SubElement(route_point, "waypoint-identifier").text = row['Description']
                    ET.SubElement(route_point, "waypoint-type").text = "USER WAYPOINT"
                rough_string = ET.tostring(flight_plan, encoding='utf-8', method='xml')
                reparsed = minidom.parseString(rough_string)
                xml_string = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + reparsed.toprettyxml(indent="  ")[23:]
                data = xml_string
                mime = "application/xml"
                ext = ".fpl"
            
            elif output_format == 'KML':
                # 生成KML
                kml = ET.Element("kml", xmlns="http://www.opengis.net/kml/2.2")
                document = ET.SubElement(kml, "Document")
                ET.SubElement(document, "name").text = sheet_name
                for _, row in chunk_df.iterrows():
                    placemark = ET.SubElement(document, "Placemark")
                    ET.SubElement(placemark, "name").text = row['Description']
                    point = ET.SubElement(placemark, "Point")
                    ET.SubElement(point, "coordinates").text = f"{row['Longitude']:.4f},{row['Latitude']:.4f},0"
                rough_string = ET.tostring(kml, encoding='utf-8', method='xml')
                reparsed = minidom.parseString(rough_string)
                xml_string = '<?xml version="1.0" encoding="UTF-8"?>\n' + reparsed.toprettyxml(indent="  ")[23:]
                data = xml_string
                mime = "application/vnd.google-earth.kml+xml"
                ext = ".kml"
            
            elif output_format == 'GPX':
                # 生成GPX (用route)
                gpx = ET.Element("gpx", version="1.1", creator="Streamlit Converter", xmlns="http://www.topografix.com/GPX/1/1")
                rte = ET.SubElement(gpx, "rte")
                ET.SubElement(rte, "name").text = sheet_name
                for _, row in chunk_df.iterrows():
                    rtept = ET.SubElement(rte, "rtept", lat=f"{row['Latitude']:.4f}", lon=f"{row['Longitude']:.4f}")
                    ET.SubElement(rtept, "name").text = row['Description']
                rough_string = ET.tostring(gpx, encoding='utf-8', method='xml')
                reparsed = minidom.parseString(rough_string)
                xml_string = '<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n' + reparsed.toprettyxml(indent="  ")[23:]
                data = xml_string
                mime = "application/gpx+xml"
                ext = ".gpx"
            
            elif output_format == 'Excel':
                # 生成Excel，根据coord_format格式化
                output_df = chunk_df.copy()
                if coord_format == "Degrees Minutes Seconds (DMS)":
                    output_df['Latitude'] = output_df['Latitude'].apply(lambda x: dd_to_dms(x, is_lat=True))
                    output_df['Longitude'] = output_df['Longitude'].apply(lambda x: dd_to_dms(x, is_lat=False))
                elif coord_format == "Degrees Minutes (DM)":
                    output_df['Latitude'] = output_df['Latitude'].apply(lambda x: dd_to_dm(x, is_lat=True))
                    output_df['Longitude'] = output_df['Longitude'].apply(lambda x: dd_to_dm(x, is_lat=False))
                # else: 保持Decimal
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    output_df.to_excel(writer, index=False)
                data = output.getvalue()
                mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ext = ".xlsx"
            
            # 下载按钮
            chunk_num = '_' + str(i//max_waypoints + 1) if i > 0 else ''
            file_name = f"{base_filename}{chunk_num}{ext}"
            st.download_button(
                label=f"Download {file_name}",
                data=data,
                file_name=file_name,
                mime=mime
            )
            st.success(f"{file_name} is ready to download! Enjoy~")
