import streamlit as st
import pandas as pd
import io
import re
import matplotlib.pyplot as plt
import datetime
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="Rider POD & Idle Time Analysis", layout="centered")

st.title("🚚 Rider POD & Idle Time Analysis Web App")

st.markdown("""
This tool lets you upload Detrack Excel files and get:
- Rider **POD tracking summary + charts (POD count & weight)**
- Rider **idle time, mileage, and max speed summary + charts**
- Downloadable tables and **charts inside downloadable Excel files**
- All data restricted to working hours: 8:30 AM – 5:30 PM

---
""")

# ----------------------------- POD Section -----------------------------
st.header("📦 POD Tracking Summary")
st.markdown("ℹ️ **Go to Detrack > Today's Job > All Jobs > Export Excel.**")
pod_file = st.file_uploader("Upload POD Excel file (delivery item)", type=["xlsx", "xls"], key="pod")

if pod_file:
    df_pod = pd.read_excel(pod_file)
    df_pod.columns = df_pod.columns.str.strip()

    st.success("✅ POD file uploaded successfully!")

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

        # POD Count Chart
        pod_summary_sorted = pod_summary.sort_values("Total_PODs", ascending=False)
        fig_pod, ax_pod = plt.subplots(figsize=(8, 5))
        bars_pod = ax_pod.bar(pod_summary_sorted["Assign to"], pod_summary_sorted["Total_PODs"], color="orange")
        ax_pod.set_title("Total PODs per Rider")
        ax_pod.set_xlabel("Rider")
        ax_pod.set_ylabel("Total PODs")
        plt.xticks(rotation=60, ha='right')
        for bar in bars_pod:
            height = bar.get_height()
            ax_pod.annotate(f"{height}", xy=(bar.get_x() + bar.get_width() / 2, height), xytext=(0, 3), textcoords="offset points", ha='center', va='bottom')
        st.pyplot(fig_pod)

        pod_img_buf = io.BytesIO()
        fig_pod.savefig(pod_img_buf, format='png', bbox_inches="tight")
        pod_img_buf.seek(0)
        st.download_button("⬇️ Download POD Chart (PNG)", pod_img_buf, "pod_chart.png", "image/png")

        # Weight Chart
        pod_summary_sorted_weight = pod_summary.sort_values("Total_Weight", ascending=False)
        fig_weight, ax_weight = plt.subplots(figsize=(8, 5))
        bars_weight = ax_weight.bar(pod_summary_sorted_weight["Assign to"], pod_summary_sorted_weight["Total_Weight"], color="blue")
        ax_weight.set_title("Total Weight per Rider (kg)")
        ax_weight.set_xlabel("Rider")
        ax_weight.set_ylabel("Total Weight")
        plt.xticks(rotation=60, ha='right')
        for bar in bars_weight:
            height = bar.get_height()
            ax_weight.annotate(f"{height:.1f}", xy=(bar.get_x() + bar.get_width() / 2, height), xytext=(0, 3), textcoords="offset points", ha='center', va='bottom')
        st.pyplot(fig_weight)

        # Save Excel with charts
        output_pod = io.BytesIO()
        with pd.ExcelWriter(output_pod, engine='openpyxl') as writer:
            pod_summary.to_excel(writer, index=False, sheet_name="POD Summary")
            workbook = writer.book
            worksheet = writer.sheets["POD Summary"]

            # Charts
            pod_img_buf.seek(0)
            img1 = XLImage(pod_img_buf)
            img1.anchor = "G1"

            weight_img_buf = io.BytesIO()
            fig_weight.savefig(weight_img_buf, format='png', bbox_inches="tight")
            weight_img_buf.seek(0)
            img2 = XLImage(weight_img_buf)
            img2.anchor = "G32"

            worksheet.add_image(img1)
            worksheet.add_image(img2)

        processed_pod = output_pod.getvalue()
        file_name_pod = f"pod_summary_{delivery_date}.xlsx"
        st.download_button("⬇️ Download POD Summary Excel with charts", processed_pod, file_name_pod, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ----------------------------- Idle Section -----------------------------
st.header("🕒 Idle Time, Mileage & Max Speed Analysis")
st.markdown("ℹ️ **Go to Detrack > Vehicles > Click on download route for each individual riders and upload all here.**")
rider_files = st.file_uploader("Upload multiple rider Excel files (vehicle route)", type=["xlsx", "xls"], accept_multiple_files=True, key="idle")

if rider_files:
    summary = []

    for file in rider_files:
        date_match = re.search(r'\d{4}-\d{2}-\d{2}', file.name)
        date_str = date_match.group(0) if date_match else "unknown_date"

        xl = pd.ExcelFile(file)
        rider_name = xl.sheet_names[0]
        df = pd.read_excel(file, sheet_name=rider_name)

        if "Time" not in df.columns or "Mileage (km)" not in df.columns or "Speed (km/h)" not in df.columns:
            st.error(f"❌ File {file.name} is missing required columns.")
            continue

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
            status = "Not working for the day"
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
            status = "Working for the day"

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

        def format_hours_mins(x):
            if x == 0 or pd.isna(x):
                return "0 hr 0 min"
            hours = int(x // 60)
            mins = int(x % 60)
            return f"{hours} hr {mins} min"

        summary_df["Idle >15 mins (formatted)"] = summary_df["Idle time >15 mins (mins)"].apply(format_hours_mins)
        summary_df["Idle time >15 mins (hrs)"] = summary_df["Idle time >15 mins (mins)"] / 60

        # Charts
        summary_df_sorted_idle = summary_df.sort_values("Idle time >15 mins (hrs)", ascending=False)
        fig_idle, ax_idle = plt.subplots(figsize=(8, 5))
        bars_idle = ax_idle.bar(summary_df_sorted_idle["Rider"], summary_df_sorted_idle["Idle time >15 mins (hrs)"], color="skyblue")
        ax_idle.set_title("Idle Time >15 mins per Rider (hours)")
        ax_idle.set_xlabel("Rider")
        ax_idle.set_ylabel("Idle Time (hrs)")
        plt.xticks(rotation=60, ha='right')
        for bar in bars_idle:
            height = bar.get_height()
            ax_idle.annotate(f"{height:.1f}", xy=(bar.get_x() + bar.get_width() / 2, height), xytext=(0, 3), textcoords="offset points", ha='center', va='bottom')

        summary_df_sorted_speed = summary_df.sort_values("Max speed (km/h)", ascending=False)
        fig_speed, ax_speed = plt.subplots(figsize=(8, 5))
        bars_speed = ax_speed.bar(summary_df_sorted_speed["Rider"], summary_df_sorted_speed["Max speed (km/h)"], color="green")
        ax_speed.set_title("Max Speed per Rider (km/h)")
        ax_speed.set_xlabel("Rider")
        ax_speed.set_ylabel("Max Speed (km/h)")
        plt.xticks(rotation=60, ha='right')
        for bar in bars_speed:
            height = bar.get_height()
            ax_speed.annotate(f"{height:.0f}", xy=(bar.get_x() + bar.get_width() / 2, height), xytext=(0, 3), textcoords="offset points", ha='center', va='bottom')

        summary_df_sorted_mileage = summary_df.sort_values("Total mileage (km)", ascending=False)
        fig_mileage, ax_mileage = plt.subplots(figsize=(8, 5))
        bars_mileage = ax_mileage.bar(summary_df_sorted_mileage["Rider"], summary_df_sorted_mileage["Total mileage (km)"], color="purple")
        ax_mileage.set_title("Total Mileage per Rider (km)")
        ax_mileage.set_xlabel("Rider")
        ax_mileage.set_ylabel("Total Mileage (km)")
        plt.xticks(rotation=60, ha='right')
        for bar in bars_mileage:
            height = bar.get_height()
            ax_mileage.annotate(f"{height:.1f}", xy=(bar.get_x() + bar.get_width() / 2, height), xytext=(0, 3), textcoords="offset points", ha='center', va='bottom', fontsize=6)

        # Show on app and PNG download
        st.pyplot(fig_idle)
        idle_img_buf = io.BytesIO()
        fig_idle.savefig(idle_img_buf, format='png', bbox_inches="tight")
        idle_img_buf.seek(0)
        st.download_button("⬇️ Download Idle Time Chart (PNG)", idle_img_buf, "idle_time_chart.png", "image/png")

        st.pyplot(fig_speed)
        speed_img_buf = io.BytesIO()
        fig_speed.savefig(speed_img_buf, format='png', bbox_inches="tight")
        speed_img_buf.seek(0)
        st.download_button("⬇️ Download Max Speed Chart (PNG)", speed_img_buf, "max_speed_chart.png", "image/png")

        st.pyplot(fig_mileage)
        mileage_img_buf = io.BytesIO()
        fig_mileage.savefig(mileage_img_buf, format='png', bbox_inches="tight")
        mileage_img_buf.seek(0)
        st.download_button("⬇️ Download Mileage Chart (PNG)", mileage_img_buf, "mileage_chart.png", "image/png")

        summary_df = summary_df.drop(columns=["Idle time >15 mins (hrs)"])

        st.subheader("📄 Idle Time Summary Table")
        st.dataframe(summary_df)

        # Save Excel with charts
        output_idle = io.BytesIO()
        with pd.ExcelWriter(output_idle, engine='openpyxl') as writer:
            summary_df.to_excel(writer, index=False, sheet_name="Idle Summary")
            workbook = writer.book
            worksheet = writer.sheets["Idle Summary"]

            idle_img_buf.seek(0)
            img1 = XLImage(idle_img_buf)
            img1.anchor = "M1"

            speed_img_buf.seek(0)
            img2 = XLImage(speed_img_buf)
            img2.anchor = "M32"

            mileage_img_buf.seek(0)
            img3 = XLImage(mileage_img_buf)
            img3.anchor = "M61"

            worksheet.add_image(img1)
            worksheet.add_image(img2)
            worksheet.add_image(img3)

        processed_idle = output_idle.getvalue()
        output_date = summary[0]["Date"] if summary else "unknown_date"
        file_name_idle = f"idle_summary_{output_date}.xlsx"
        st.download_button("⬇️ Download Idle Time Summary Excel with charts", processed_idle, file_name_idle, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# -----------------------------
# Section 3: Cartrack Summary (final cleaned version)
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
        df_trip["Registration"] = registration

        df_trip["Trip Distance"] = pd.to_numeric(df_trip.get("Trip Distance", 0), errors="coerce").fillna(0)
        speed_col = next((c for c in df_trip.columns if re.match(r"Speeding", c, re.IGNORECASE)), None)
        df_trip["Speeding_Count"] = pd.to_numeric(df_trip.get(speed_col, 0), errors="coerce").fillna(0) if speed_col else 0

        # --- Read Fuel Efficiency data ---
        xl_fuel = pd.ExcelFile(fuel_file)
        raw_fuel = xl_fuel.parse(xl_fuel.sheet_names[0], header=None)
        header_idx = raw_fuel[raw_fuel.iloc[:, 0].astype(str).str.contains("Vehicle Registration", na=False)].index[0]
        df_fuel = xl_fuel.parse(xl_fuel.sheet_names[0], skiprows=header_idx)
        df_fuel.columns = df_fuel.columns.str.strip()
        df_fuel.rename(columns={"Vehicle Registration": "Registration"}, inplace=True)

        # 💡 Super safe cleaning fix
        df_fuel["Registration"] = df_fuel["Registration"].astype(str).str.strip()
        df_fuel["Registration"] = df_fuel["Registration"].replace("nan", "")

        fuel_col = next((c for c in df_fuel.columns if "Fuel Consumed" in c), None)
        dist_col = next((c for c in df_fuel.columns if "Distance Travelled" in c), None)

        if fuel_col and fuel_col in df_fuel.columns:
            df_fuel["Fuel Consumed (litres)"] = pd.to_numeric(df_fuel[fuel_col], errors="coerce").fillna(0)
        else:
            df_fuel["Fuel Consumed (litres)"] = 0

        if dist_col and dist_col in df_fuel.columns:
            df_fuel["Distance Travelled (km)"] = pd.to_numeric(df_fuel[dist_col], errors="coerce").fillna(0)
        else:
            df_fuel["Distance Travelled (km)"] = 0

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

        # --- Assign drivers using df_trip first, then fallback ---
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
            if isinstance(td, str) and td.strip():
                return td
            reg = row.get("Registration", "")
            if reg in override_map:
                return override_map[reg]
            loc = str(row.get("End Location", "") or "")
            if re.search(r"Punggol|Hougang", loc, re.IGNORECASE):
                return "Abdul Rahman"
            if re.search(r"Woodlands|Yishun|Jurong East", loc, re.IGNORECASE):
                return "Sugathan"
            if re.search(r"Changi South", loc, re.IGNORECASE):
                return "Mohd"
            if re.search(r"Pasir Panjang", loc, re.IGNORECASE):
                return "Toh"
            if re.search(r"Kallang", loc, re.IGNORECASE) and not re.search(r"Pasir Panjang", loc, re.IGNORECASE):
                return "Masari"
            return "Unknown"

        reg_info = (
            df_context.groupby("Registration", as_index=False)
                      .agg(
                          TripDriver=("TripDriver", lambda x: next((v for v in x if isinstance(v, str) and v.strip()), "")),
                          EndLocation=("End Location", lambda x: next((v for v in x if isinstance(v, str) and v.strip()), ""))
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
            lambda r: r["Total_Mileage_km"]/r["Total_Fuel_Litres"] if r["Total_Fuel_Litres"] > 0 else None,
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
