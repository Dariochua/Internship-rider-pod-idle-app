import streamlit as st
import pandas as pd
import io
import re
import matplotlib.pyplot as plt
import datetime
import difflib

# NOTE: xlsxwriter is the engine we’re using for embedded charts
import xlsxwriter

st.set_page_config(page_title="Rider POD & Idle Time Analysis", layout="centered")

st.title("🚚 Rider POD & Idle Time Analysis Web App")

st.markdown("""
This tool lets you upload Detrack Excel files and get:
- Rider **POD tracking summary + charts (POD count & weight)**
- Rider **idle time, mileage, and max speed summary + charts**
- Downloadable tables and **downloadable charts embedded in Excel**
- All data restricted to working hours: 8:30 AM – 5:30 PM

---
""")

# -----------------------------
# Section 1: POD Tracking
# -----------------------------
st.header("📦 POD Tracking Summary")

pod_file = st.file_uploader("Upload POD Excel file (delivery item)", type=["xlsx", "xls"], key="pod")

if pod_file:
    df_pod = pd.read_excel(pod_file)
    df_pod.columns = df_pod.columns.str.strip()

    st.success("✅ POD file uploaded successfully!")

    if {"Assign to","POD Time","Weight","Delivery Date"}.issubset(df_pod.columns):
        try:
            df_pod["Delivery Date"] = pd.to_datetime(df_pod["Delivery Date"], errors='coerce')
            df_pod["POD Time"] = pd.to_datetime(df_pod["POD Time"], errors='coerce').dt.time

            df_pod["POD DateTime"] = df_pod.apply(
                lambda r: datetime.datetime.combine(r["Delivery Date"], r["POD Time"])
                          if pd.notnull(r["Delivery Date"]) and pd.notnull(r["POD Time"]) else pd.NaT,
                axis=1
            )

            delivery_date_mode = df_pod["Delivery Date"].mode()[0]
            delivery_date = delivery_date_mode.strftime("%Y-%m-%d")
        except:
            delivery_date = "unknown_date"
    else:
        st.error("❌ Required columns 'Assign to', 'POD Time', 'Weight', or 'Delivery Date' not found.")
        delivery_date = "unknown_date"

    if delivery_date != "unknown_date":
        pod_summary = df_pod.groupby("Assign to").agg(
            Earliest_POD=("POD DateTime", "min"),
            Latest_POD=("POD DateTime", "max"),
            Total_PODs=("POD DateTime", "count"),
            Total_Weight=("Weight", "sum")
        ).reset_index()

        st.subheader("📄 POD Summary Table")
        st.dataframe(pod_summary)

        # ----- Create Excel with embedded charts -----
        output_pod = io.BytesIO()
        with pd.ExcelWriter(output_pod, engine='xlsxwriter') as writer:
            pod_summary.to_excel(writer, index=False, sheet_name='POD Summary')
            workbook  = writer.book
            worksheet = writer.sheets['POD Summary']
            n = len(pod_summary)

            # Chart 1: Total PODs
            chart1 = workbook.add_chart({'type': 'column'})
            chart1.add_series({
                'name':       'Total PODs',
                'categories': ['POD Summary', 1, 0, n, 0],
                'values':     ['POD Summary', 1, 2, n, 2],
            })
            chart1.set_title({'name': 'Total PODs per Rider'})
            chart1.set_x_axis({'name': 'Rider'})
            chart1.set_y_axis({'name': 'POD Count'})
            worksheet.insert_chart('F2', chart1, {'x_scale': 1.5, 'y_scale': 1.5})

            # Chart 2: Total Weight
            chart2 = workbook.add_chart({'type': 'column'})
            chart2.add_series({
                'name':       'Total Weight',
                'categories': ['POD Summary', 1, 0, n, 0],
                'values':     ['POD Summary', 1, 3, n, 3],
            })
            chart2.set_title({'name': 'Total Weight per Rider'})
            chart2.set_x_axis({'name': 'Rider'})
            chart2.set_y_axis({'name': 'Weight'})
            worksheet.insert_chart('F20', chart2, {'x_scale': 1.5, 'y_scale': 1.5})

        processed_pod = output_pod.getvalue()
        st.download_button(
            "⬇️ Download POD Summary with Charts (Excel)",
            processed_pod,
            f"pod_summary_{delivery_date}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# -----------------------------
# Section 2: Idle Time Analysis
# -----------------------------
st.header("🕒 Idle Time, Mileage & Max Speed Analysis")

rider_files = st.file_uploader(
    "Upload multiple rider Excel files (vehicle route)",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
    key="idle"
)

if rider_files:
    summary = []
    for file in rider_files:
        date_match = re.search(r'\d{4}-\d{2}-\d{2}', file.name)
        date_str = date_match.group(0) if date_match else "unknown_date"

        xl = pd.ExcelFile(file)
        rider_name = xl.sheet_names[0]
        df = pd.read_excel(file, sheet_name=rider_name)

        if {"Time","Mileage (km)","Speed (km/h)"}.difference(df.columns):
            st.error(f"❌ File {file.name} is missing required columns.")
            continue

        df['Time'] = pd.to_datetime(df['Time'], format='%I:%M:%S %p', errors='coerce')
        df['Idle'] = df['Mileage (km)'] == 0

        work_start = datetime.time(8, 30)
        work_end   = datetime.time(17, 30)
        df['Time_only'] = df['Time'].dt.time
        df_working = df[(df['Time_only'] >= work_start) & (df['Time_only'] <= work_end)]

        total_mileage = df_working['Mileage (km)'].sum()
        if total_mileage == 0:
            total_idle = total_over_15 = num_over_15 = max_speed = 0
            status = "Not working for the day"
        else:
            # compute idle periods
            idle_periods = []
            current_start = None
            for _, row in df.iterrows():
                t = row['Time'].time()
                if t < work_start or t > work_end:
                    if current_start is not None:
                        idle_periods.append((current_start, row['Time']))
                        current_start = None
                    continue

                if row['Idle']:
                    if current_start is None:
                        current_start = row['Time']
                else:
                    if current_start is not None:
                        idle_periods.append((current_start, row['Time']))
                        current_start = None

            if current_start is not None:
                idle_periods.append((current_start, df['Time'].iloc[-1]))

            idle_durations = [(e - s).total_seconds()/60 for s,e in idle_periods]
            total_idle      = sum(idle_durations)
            over_15         = [d for d in idle_durations if d > 15]
            total_over_15   = sum(over_15)
            num_over_15     = len(over_15)
            max_speed       = df_working['Speed (km/h)'].max()
            status          = "Working for the day"

        summary.append({
            "File": file.name,
            "Rider": rider_name,
            "Date": date_str,
            "Total idle time (mins)": total_idle,
            "Idle time >15 mins (mins)": total_over_15,
            "Num idle periods >15 mins": num_over_15,
            "Total mileage (km)": total_mileage,
            "Max speed (km/h)": max_speed,
            "Status": status
        })

    if summary:
        summary_df = pd.DataFrame(summary)

        # add a column for the chart data
        summary_df["Idle time >15 mins (hrs)"] = summary_df["Idle time >15 mins (mins)"] / 60

        st.subheader("📄 Idle Time Summary Table")
        st.dataframe(summary_df)

        # ----- Create Excel with embedded charts -----
        output_idle = io.BytesIO()
        with pd.ExcelWriter(output_idle, engine='xlsxwriter') as writer:
            summary_df.to_excel(writer, index=False, sheet_name='Idle Summary')
            workbook = writer.book
            ws       = writer.sheets['Idle Summary']
            n = len(summary_df)

            # Idle >15 hrs chart
            c1 = workbook.add_chart({'type': 'column'})
            c1.add_series({
                'name':       'Idle >15 mins (hrs)',
                'categories': ['Idle Summary', 1, 1, n, 1],  # Rider names in col B
                'values':     ['Idle Summary', 1, summary_df.columns.get_loc('Idle time >15 mins (hrs)'), n, summary_df.columns.get_loc('Idle time >15 mins (hrs)')],
            })
            c1.set_title({'name': 'Idle >15 mins per Rider (hrs)'})
            ws.insert_chart('K2', c1, {'x_scale': 1.3, 'y_scale': 1.3})

            # Max Speed chart
            c2 = workbook.add_chart({'type': 'column'})
            c2.add_series({
                'name':       'Max Speed (km/h)',
                'categories': ['Idle Summary', 1, 1, n, 1],
                'values':     ['Idle Summary', 1, summary_df.columns.get_loc('Max speed (km/h)'), n, summary_df.columns.get_loc('Max speed (km/h)')],
            })
            c2.set_title({'name': 'Max Speed per Rider'})
            ws.insert_chart('K20', c2, {'x_scale': 1.3, 'y_scale': 1.3})

            # Total Mileage chart
            c3 = workbook.add_chart({'type': 'column'})
            c3.add_series({
                'name':       'Total Mileage (km)',
                'categories': ['Idle Summary', 1, 1, n, 1],
                'values':     ['Idle Summary', 1, summary_df.columns.get_loc('Total mileage (km)'), n, summary_df.columns.get_loc('Total mileage (km)')],
            })
            c3.set_title({'name': 'Total Mileage per Rider'})
            ws.insert_chart('K38', c3, {'x_scale': 1.3, 'y_scale': 1.3})

        processed_idle = output_idle.getvalue()
        out_date = summary_df.loc[0, "Date"]
        st.download_button(
            "⬇️ Download Idle Summary with Charts (Excel)",
            processed_idle,
            f"idle_time_summary_{out_date}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# -----------------------------
# Section 3: Cartrack Summary
# -----------------------------
st.header("🚗 Cartrack Summary")
trip_file = st.file_uploader("Upload Summary Trip Report", type=["xlsx", "xls"], key="trip")
fuel_file = st.file_uploader("Upload Fuel Efficiency Report", type=["xlsx", "xls"], key="fuel")

if trip_file and fuel_file:
    try:
        # --- your existing Cartrack logic unchanged ---
        xl_trip = pd.ExcelFile(trip_file)
        meta    = xl_trip.parse(xl_trip.sheet_names[0], header=None, nrows=15)
        reg_row = meta[meta.iloc[:,0].astype(str).str.contains("Registration",na=False)]
        registration = str(reg_row.iloc[0,1]).strip() if not reg_row.empty else ""

        raw_trip = xl_trip.parse(xl_trip.sheet_names[0], header=None)
        start_idx = raw_trip[raw_trip.iloc[:,0]=="Driver"].index[0]
        df_trip = xl_trip.parse(xl_trip.sheet_names[0], skiprows=start_idx)
        df_trip.columns = df_trip.columns.str.strip()
        df_trip.rename(columns={"Driver":"TripDriver"}, inplace=True)
        df_trip["Registration"] = registration
        df_trip["Trip Distance"] = pd.to_numeric(df_trip.get("Trip Distance",0),errors="coerce").fillna(0)
        speed_col = next((c for c in df_trip.columns if re.match(r"Speeding",c,re.IGNORECASE)),None)
        df_trip["Speeding_Count"] = pd.to_numeric(df_trip.get(speed_col,0),errors="coerce").fillna(0) if speed_col else 0

        xl_fuel = pd.ExcelFile(fuel_file)
        raw_fuel = xl_fuel.parse(xl_fuel.sheet_names[0], header=None)
        header_idx = raw_fuel[raw_fuel.iloc[:,0].astype(str).str.contains("Vehicle Registration",na=False)].index[0]
        df_fuel = xl_fuel.parse(xl_fuel.sheet_names[0], skiprows=header_idx)
        df_fuel.columns = df_fuel.columns.str.strip()
        df_fuel.rename(columns={"Vehicle Registration":"Registration"}, inplace=True)
        df_fuel["Registration"] = df_fuel["Registration"].astype(str).str.strip()
        fuel_col = next((c for c in df_fuel.columns if re.match(r"Fuel Consumed",c,re.IGNORECASE)),None)
        dist_col = next((c for c in df_fuel.columns if re.match(r"Distance Travelled",c,re.IGNORECASE)),None)
        df_fuel["Fuel Consumed (litres)"] = pd.to_numeric(df_fuel.get(fuel_col,0),errors="coerce").fillna(0)
        df_fuel["Distance Travelled (km)"] = pd.to_numeric(df_fuel.get(dist_col,0),errors="coerce").fillna(0)

        trip_agg = df_trip.groupby("Registration",as_index=False).agg(
            Total_Trip_Distance_km=("Trip Distance","sum"),
            Total_Speeding_Count=("Speeding_Count","sum")
        )
        fuel_agg = df_fuel.groupby("Registration",as_index=False).agg(
            Total_Fuel_Litres=("Fuel Consumed (litres)","sum"),
            Total_Distance_Travelled_km=("Distance Travelled (km)","sum")
        )
        df_agg = pd.merge(trip_agg,fuel_agg,on="Registration",how="outer").fillna(0)

        override_map = {
            "GBB933E":"Abdul Rahman","GBB933Z":"Mohd","GBC8305D":"Sugathan",
            "GBC9338C":"Toh","GX9339E":"Masari","GY933T":"Mohd Hairul","GBB933X":"Unknown"
        }
        meta2 = pd.merge(df_trip, df_fuel[["Registration"]], on="Registration", how="outer")
        def assign_driver(r):
            if isinstance(r.TripDriver,str) and r.TripDriver.strip(): return r.TripDriver
            reg = r.Registration
            if reg in override_map: return override_map[reg]
            loc = str(r.get("End Location",""))
            if re.search(r"Punggol|Hougang",loc,re.IGNORECASE): return "Abdul Rahman"
            if re.search(r"Woodlands|Yishun|Jurong East",loc,re.IGNORECASE): return "Sugathan"
            if re.search(r"Changi South",loc,re.IGNORECASE): return "Mohd"
            if re.search(r"Pasir Panjang",loc,re.IGNORECASE): return "Toh"
            if re.search(r"Kallang",loc,re.IGNORECASE) and not re.search(r"Pasir Panjang",loc,re.IGNORECASE): return "Masari"
            return "Unknown"

        reg_info = (
            meta2.groupby("Registration",as_index=False)
                 .agg(
                     TripDriver=("TripDriver",lambda x: next((v for v in x if isinstance(v,str) and v), "")),
                     EndLocation=("End Location",lambda x: next((v for v in x if isinstance(v,str) and v), ""))
                 )
        )
        reg_info["Driver"] = reg_info.apply(assign_driver,axis=1)

        st.subheader("📝 Rider ↔ Vehicle Mapping")
        st.dataframe(reg_info[["Registration","Driver"]].drop_duplicates().sort_values(["Driver","Registration"]))

        df_summary = pd.merge(reg_info[["Registration","Driver"]], df_agg, on="Registration", how="left").fillna(0)
        summary2 = df_summary.groupby("Driver",as_index=False).agg(
            Total_Speeding_Count=("Total_Speeding_Count","sum"),
            Total_Mileage_km=("Total_Trip_Distance_km","sum"),
            Total_Fuel_Litres=("Total_Fuel_Litres","sum")
        )
        summary2["Fuel_Efficiency (km/l)"] = summary2.apply(
            lambda r: r.Total_Mileage_km/r.Total_Fuel_Litres if r.Total_Fuel_Litres>0 else None, axis=1
        )
        st.subheader("📄 Rider Fuel & Mileage Summary")
        st.dataframe(summary2)

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as writer:
            reg_info[["Registration","Driver"]].drop_duplicates().to_excel(writer, index=False, sheet_name="Mapping")
            summary2.to_excel(writer, index=False, sheet_name="Summary")
        st.download_button(
            "⬇️ Download Cartrack Summary Excel",
            buf.getvalue(),
            "cartrack_summary.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"❌ Processing error: {e}")
