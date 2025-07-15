import streamlit as st
import pandas as pd
import io
import re
import matplotlib.pyplot as plt
import datetime
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="Rider POD & Idle Time Analysis", layout="centered")

st.title("🚚 Rider POD & Idle Time Analysis Web App")

st.markdown("""
This tool lets you upload Detrack Excel files and get:
- Rider **POD tracking summary + charts (POD count & weight)**
- Rider **idle time, mileage, and max speed summary + charts**
- Downloadable tables and **downloadable charts + charts included in Excel**
""")

# -----------------------------
# Section 1: POD Tracking
# -----------------------------
st.header("📦 POD Tracking Summary")

pod_file = st.file_uploader("Upload POD Excel file", type=["xlsx", "xls"], key="pod")

if pod_file:
    df_pod = pd.read_excel(pod_file)
    df_pod.columns = df_pod.columns.str.strip()

    if "POD Time" in df_pod.columns and "Assign to" in df_pod.columns and "Weight" in df_pod.columns and "Delivery Date" in df_pod.columns:
        df_pod["Delivery Date"] = pd.to_datetime(df_pod["Delivery Date"], errors='coerce')
        df_pod["POD Time"] = pd.to_datetime(df_pod["POD Time"], errors='coerce').dt.time
        df_pod["POD DateTime"] = df_pod.apply(
            lambda row: datetime.datetime.combine(row["Delivery Date"], row["POD Time"]) if pd.notnull(row["Delivery Date"]) and pd.notnull(row["POD Time"]) else pd.NaT,
            axis=1
        )
        delivery_date_mode = df_pod["Delivery Date"].mode()[0]
        delivery_date = delivery_date_mode.strftime("%Y-%m-%d")
    else:
        st.error("❌ Required columns not found.")
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

        # Chart for POD count
        pod_summary_sorted = pod_summary.sort_values("Total_PODs", ascending=False)
        fig_pod, ax_pod = plt.subplots(figsize=(8, 5))
        bars_pod = ax_pod.bar(pod_summary_sorted["Assign to"], pod_summary_sorted["Total_PODs"], color="orange")
        ax_pod.set_title("Total PODs per Rider")
        plt.xticks(rotation=60, ha='right')
        for bar in bars_pod:
            height = bar.get_height()
            ax_pod.annotate(f"{height}", xy=(bar.get_x() + bar.get_width() / 2, height), xytext=(0, 3), textcoords="offset points", ha='center', va='bottom')

        st.pyplot(fig_pod)

        pod_img_buf = io.BytesIO()
        fig_pod.savefig(pod_img_buf, format='png', bbox_inches="tight")
        pod_img_buf.seek(0)
        st.download_button("⬇️ Download POD Chart (PNG)", pod_img_buf, "pod_chart.png", "image/png")

        # Chart for weight
        pod_summary_sorted_weight = pod_summary.sort_values("Total_Weight", ascending=False)
        fig_weight, ax_weight = plt.subplots(figsize=(8, 5))
        bars_weight = ax_weight.bar(pod_summary_sorted_weight["Assign to"], pod_summary_sorted_weight["Total_Weight"], color="blue")
        ax_weight.set_title("Total Weight per Rider")
        plt.xticks(rotation=60, ha='right')
        for bar in bars_weight:
            height = bar.get_height()
            ax_weight.annotate(f"{height:.1f}", xy=(bar.get_x() + bar.get_width() / 2, height), xytext=(0, 3), textcoords="offset points", ha='center', va='bottom')

        st.pyplot(fig_weight)

        # Create Excel with charts
        wb = Workbook()
        ws = wb.active
        ws.title = "POD Summary"

        for r in dataframe_to_rows(pod_summary, index=False, header=True):
            ws.append(r)

        pod_img_buf.seek(0)
        pod_img = XLImage(pod_img_buf)
        pod_img.anchor = "G1"
        ws.add_image(pod_img)

        weight_img_buf = io.BytesIO()
        fig_weight.savefig(weight_img_buf, format='png', bbox_inches="tight")
        weight_img_buf.seek(0)
        weight_img = XLImage(weight_img_buf)
        weight_img.anchor = "G25"
        ws.add_image(weight_img)

        output_pod = io.BytesIO()
        wb.save(output_pod)
        processed_pod = output_pod.getvalue()

        file_name_pod = f"pod_summary_{delivery_date}.xlsx"
        st.download_button("⬇️ Download POD Summary Excel", processed_pod, file_name_pod, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# -----------------------------
# Section 2: Idle Analysis
# -----------------------------
st.header("🕒 Idle Time, Mileage & Max Speed Analysis")

rider_files = st.file_uploader("Upload multiple rider files", type=["xlsx", "xls"], accept_multiple_files=True, key="idle")

if rider_files:
    summary = []

    for file in rider_files:
        xl = pd.ExcelFile(file)
        rider_name = xl.sheet_names[0]
        df = pd.read_excel(file, sheet_name=rider_name)

        df['Time'] = pd.to_datetime(df['Time'], format='%I:%M:%S %p', errors='coerce')
        df['Idle'] = df['Mileage (km)'] == 0

        work_start = datetime.time(8, 30)
        work_end = datetime.time(17, 30)
        df['Time_only'] = df['Time'].dt.time
        df_working = df[(df['Time_only'] >= work_start) & (df['Time_only'] <= work_end)]
        total_mileage = df_working['Mileage (km)'].sum()

        if total_mileage == 0:
            total_idle = 0
            total_over_15 = 0
            num_over_15 = 0
            max_speed = 0
        else:
            idle_periods = []
            current_start = None
            for idx, row in df.iterrows():
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
            if current_start is not None and work_start <= current_start.time() <= work_end:
                idle_periods.append((current_start, df['Time'].iloc[-1]))

            idle_durations = [(end - start).total_seconds() / 60 for start, end in idle_periods]
            total_idle = sum(idle_durations)
            over_15 = [d for d in idle_durations if d > 15]
            total_over_15 = sum(over_15)
            num_over_15 = len(over_15)
            max_speed = df_working['Speed (km/h)'].max()

        summary.append({
            "File": file.name,
            "Rider": rider_name,
            "Total idle time (mins)": total_idle,
            "Idle time >15 mins (mins)": total_over_15,
            "Num idle periods >15 mins": num_over_15,
            "Total mileage (km)": total_mileage,
            "Max speed (km/h)": max_speed,
        })

    summary_df = pd.DataFrame(summary)

    def format_hours_mins(x):
        if x == 0 or pd.isna(x):
            return "0 hr 0 min"
        hours = int(x // 60)
        mins = int(x % 60)
        return f"{hours} hr {mins} min"

    summary_df["Idle >15 mins (formatted)"] = summary_df["Idle time >15 mins (mins)"].apply(format_hours_mins)
    summary_df["Idle time >15 mins (hrs)"] = summary_df["Idle time >15 mins (mins)"] / 60

    st.subheader("📄 Idle Time Summary Table")
    st.dataframe(summary_df)

    # Chart: Idle Time >15 mins
    summary_df_sorted_idle = summary_df.sort_values("Idle time >15 mins (hrs)", ascending=False)
    fig_idle, ax_idle = plt.subplots(figsize=(8, 5))
    bars_idle = ax_idle.bar(summary_df_sorted_idle["Rider"], summary_df_sorted_idle["Idle time >15 mins (hrs)"], color="skyblue")
    ax_idle.set_title("Idle Time >15 mins per Rider (hours)")
    plt.xticks(rotation=60, ha='right')
    for bar in bars_idle:
        height = bar.get_height()
        ax_idle.annotate(f"{height:.1f}", xy=(bar.get_x() + bar.get_width() / 2, height), xytext=(0, 3), textcoords="offset points", ha='center', va='bottom')
    st.pyplot(fig_idle)

    # Chart: Max Speed
    summary_df_sorted_speed = summary_df.sort_values("Max speed (km/h)", ascending=False)
    fig_speed, ax_speed = plt.subplots(figsize=(8, 5))
    bars_speed = ax_speed.bar(summary_df_sorted_speed["Rider"], summary_df_sorted_speed["Max speed (km/h)"], color="green")
    ax_speed.set_title("Max Speed per Rider (km/h)")
    plt.xticks(rotation=60, ha='right')
    for bar in bars_speed:
        height = bar.get_height()
        ax_speed.annotate(f"{height:.0f}", xy=(bar.get_x() + bar.get_width() / 2, height), xytext=(0, 3), textcoords="offset points", ha='center', va='bottom')
    st.pyplot(fig_speed)

    # Chart: Total Mileage
    summary_df_sorted_mileage = summary_df.sort_values("Total mileage (km)", ascending=False)
    fig_mileage, ax_mileage = plt.subplots(figsize=(8, 5))
    bars_mileage = ax_mileage.bar(summary_df_sorted_mileage["Rider"], summary_df_sorted_mileage["Total mileage (km)"], color="purple")
    ax_mileage.set_title("Total Mileage per Rider (km)")
    plt.xticks(rotation=60, ha='right')
    for bar in bars_mileage:
        height = bar.get_height()
        ax_mileage.annotate(f"{height:.1f}", xy=(bar.get_x() + bar.get_width() / 2, height), xytext=(0, 3), textcoords="offset points", ha='center', va='bottom')
    st.pyplot(fig_mileage)

    # Download PNG buttons
    idle_img_buf = io.BytesIO()
    fig_idle.savefig(idle_img_buf, format='png', bbox_inches="tight")
    idle_img_buf.seek(0)
    st.download_button("⬇️ Download Idle Chart (PNG)", idle_img_buf, "idle_chart.png", "image/png")

    speed_img_buf = io.BytesIO()
    fig_speed.savefig(speed_img_buf, format='png', bbox_inches="tight")
    speed_img_buf.seek(0)
    st.download_button("⬇️ Download Speed Chart (PNG)", speed_img_buf, "speed_chart.png", "image/png")

    mileage_img_buf = io.BytesIO()
    fig_mileage.savefig(mileage_img_buf, format='png', bbox_inches="tight")
    mileage_img_buf.seek(0)
    st.download_button("⬇️ Download Mileage Chart (PNG)", mileage_img_buf, "mileage_chart.png", "image/png")

    # Create Excel with charts
    wb_idle = Workbook()
    ws_idle = wb_idle.active
    ws_idle.title = "Idle Summary"

    for r in dataframe_to_rows(summary_df, index=False, header=True):
        ws_idle.append(r)

    idle_img_buf.seek(0)
    idle_img = XLImage(idle_img_buf)
    idle_img.anchor = "M1"
    ws_idle.add_image(idle_img)

    speed_img_buf.seek(0)
    speed_img = XLImage(speed_img_buf)
    speed_img.anchor = "M30"
    ws_idle.add_image(speed_img)

    mileage_img_buf.seek(0)
    mileage_img = XLImage(mileage_img_buf)
    mileage_img.anchor = "M60"
    ws_idle.add_image(mileage_img)

    output_idle = io.BytesIO()
    wb_idle.save(output_idle)
    processed_idle = output_idle.getvalue()

    st.download_button("⬇️ Download Idle Summary Excel", processed_idle, "idle_summary.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# -----------------------------
# Section 3: Cartrack Summary (Fixed Aggregation)
# -----------------------------

st.header("🚗 Cartrack Summary")

trip_file = st.file_uploader("Upload Summary Trip Report", type=["xlsx", "xls"], key="trip")
fuel_file = st.file_uploader("Upload Fuel Efficiency Report", type=["xlsx", "xls"], key="fuel")

if trip_file and fuel_file:
    try:
        # --- Read Trip Report metadata ---
        xl_trip = pd.ExcelFile(trip_file)
        meta = xl_trip.parse(xl_trip.sheet_names[0], header=None, nrows=15)
        reg_row = meta[meta.iloc[:, 0].astype(str).str.contains("Registration", na=False)]
        registration = None
        try:
            registration = str(reg_row.iloc[0, 1]).strip()
        except:
            pass

        # --- Read full trip data ---
        raw_trip = xl_trip.parse(xl_trip.sheet_names[0], header=None)
        start_idx = raw_trip[raw_trip.iloc[:, 0] == "Driver"].index[0]
        df_trip = xl_trip.parse(xl_trip.sheet_names[0], skiprows=start_idx)
        df_trip.columns = df_trip.columns.str.strip()
        df_trip.rename(columns={"Driver": "TripDriver"}, inplace=True)
        # ensure Registration column
        df_trip["Registration"] = registration
        df_trip["Trip Distance"] = pd.to_numeric(df_trip.get("Trip Distance", 0), errors="coerce").fillna(0)
        # speeding count
        speed_col = next((c for c in df_trip.columns if re.match(r"Speeding", c, re.IGNORECASE)), None)
        df_trip["Speeding_Count"] = pd.to_numeric(df_trip.get(speed_col, 0), errors="coerce").fillna(0) if speed_col else 0

        # --- Read Fuel Efficiency data ---
        xl_fuel = pd.ExcelFile(fuel_file)
        raw_fuel = xl_fuel.parse(xl_fuel.sheet_names[0], header=None)
        header_idx = raw_fuel[raw_fuel.iloc[:, 0].astype(str).str.contains("Vehicle Registration", na=False)].index[0]
        df_fuel = xl_fuel.parse(xl_fuel.sheet_names[0], skiprows=header_idx)
        df_fuel.columns = df_fuel.columns.str.strip()
        df_fuel.rename(columns={"Vehicle Registration": "Registration"}, inplace=True)
        df_fuel["Registration"] = df_fuel["Registration"].astype(str).str.strip()
        fuel_col = next((c for c in df_fuel.columns if re.match(r"Fuel Consumed", c, re.IGNORECASE)), None)
        dist_col = next((c for c in df_fuel.columns if re.match(r"Distance Travelled", c, re.IGNORECASE)), None)
        df_fuel["Fuel Consumed (litres)"] = pd.to_numeric(df_fuel.get(fuel_col, 0), errors="coerce").fillna(0)
        df_fuel["Distance Travelled (km)"] = pd.to_numeric(df_fuel.get(dist_col, 0), errors="coerce").fillna(0)

        # --- Aggregate separately ---
        trip_agg = df_trip.groupby("Registration", as_index=False).agg(
            Total_Trip_Distance_km=("Trip Distance", "sum"),
            Total_Speeding_Count=("Speeding_Count", "sum")
        )
        fuel_agg = df_fuel.groupby("Registration", as_index=False).agg(
            Total_Fuel_Litres=("Fuel Consumed (litres)", "sum"),
            Total_Distance_Travelled_km=("Distance Travelled (km)", "sum")
        )
        df_agg = pd.merge(trip_agg, fuel_agg, on="Registration", how="outer").fillna(0)

        # --- Assign drivers using df_trip context ---
        # Combine df_trip and df_fuel to get End Location context
        df_context = pd.merge(
            df_trip,
            df_fuel[["Registration"]],
            on="Registration",
            how="outer"
        )
        override_map = {
            "GBB933E": "Abdul Rahman", "GBB933Z": "Mohd", "GBC8305D": "Sugathan",
            "GBC9338C": "Toh", "GX9339E": "Masari", "GY933T": "Mohd Hairul", "GBB933X": "Unknown"
        }
        def assign_driver(row):
            td = row.get("TripDriver")
            if isinstance(td, str) and td.strip(): return td
            reg = row.get("Registration", "")
            if reg in override_map: return override_map[reg]
            loc = str(row.get("End Location", "") or "")
            if re.search(r"Punggol|Hougang", loc, re.IGNORECASE): return "Abdul Rahman"
            if re.search(r"Woodlands|Yishun|Jurong East", loc, re.IGNORECASE): return "Sugathan"
            if re.search(r"Changi South", loc, re.IGNORECASE): return "Mohd"
            if re.search(r"Pasir Panjang", loc, re.IGNORECASE): return "Toh"
            if re.search(r"Kallang", loc, re.IGNORECASE) and not re.search(r"Pasir Panjang", loc, re.IGNORECASE): return "Masari"
            return "Unknown"
        reg_info = (
            df_context.groupby("Registration", as_index=False)
                      .agg(
                          TripDriver=("TripDriver", lambda x: next((v for v in x if isinstance(v, str) and v), "")),
                          EndLocation=("End Location", lambda x: next((v for v in x if isinstance(v, str) and v), ""))
                      )
        )
        reg_info["Driver"] = reg_info.apply(assign_driver, axis=1)

        # --- Mapping table ---
        mapping = reg_info[["Registration", "Driver"]].drop_duplicates().sort_values(["Driver", "Registration"])
        st.subheader("📝 Rider ↔ Vehicle Mapping")
        st.dataframe(mapping)

        # --- Final summary by driver ---
        df_summary = pd.merge(mapping, df_agg, on="Registration", how="left").fillna(0)
        summary = df_summary.groupby("Driver", as_index=False).agg(
            Total_Speeding_Count=("Total_Speeding_Count", "sum"),
            Total_Mileage_km=("Total_Trip_Distance_km", "sum"),
            Total_Fuel_Litres=("Total_Fuel_Litres", "sum")
        )
        summary["Fuel_Efficiency (km/l)"] = summary.apply(
            lambda r: r["Total_Mileage_km"]/r["Total_Fuel_Litres"] if r["Total_Fuel_Litres"]>0 else None,
            axis=1
        )
        st.subheader("📄 Rider Fuel & Mileage Summary")
        st.dataframe(summary)

        # --- Downloadable report ---
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            mapping.to_excel(writer, index=False, sheet_name="Mapping")
            summary.to_excel(writer, index=False, sheet_name="Summary")
        st.download_button(
            "⬇️ Download Cartrack Summary Excel",
            buf.getvalue(),
            "cartrack_summary.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Processing error: {e}")
