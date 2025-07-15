import streamlit as st
import pandas as pd
import io
import re
import matplotlib.pyplot as plt
import datetime

st.set_page_config(page_title="Rider POD & Idle Time Analysis", layout="centered")

st.title("🚚 Rider POD & Idle Time Analysis Web App")

st.markdown("""
This tool lets you upload Detrack Excel files and get:
- Rider **POD tracking summary + charts (POD count & weight)**
- Rider **idle time, mileage, and max speed summary + charts**
- Downloadable tables and **downloadable charts**
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
    df_pod.columns = df_pod.columns.str.strip()  # Clean header spaces

    st.success("✅ POD file uploaded successfully!")

    if "POD Time" in df_pod.columns and "Assign to" in df_pod.columns and "Weight" in df_pod.columns and "Delivery Date" in df_pod.columns:
        try:
            df_pod["Delivery Date"] = pd.to_datetime(df_pod["Delivery Date"], errors='coerce')
            df_pod["POD Time"] = pd.to_datetime(df_pod["POD Time"], errors='coerce').dt.time

            # Combine Delivery Date and POD Time
            df_pod["POD DateTime"] = df_pod.apply(
                lambda row: datetime.datetime.combine(row["Delivery Date"], row["POD Time"]) if pd.notnull(row["Delivery Date"]) and pd.notnull(row["POD Time"]) else pd.NaT,
                axis=1
            )

            # Debug preview (optional)
            # st.write(df_pod[["Assign to", "Delivery Date", "POD Time", "POD DateTime"]].head(50))

            # Use most common delivery date for file name
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

        # POD count chart
        pod_summary_sorted = pod_summary.sort_values("Total_PODs", ascending=False)
        fig_pod, ax_pod = plt.subplots(figsize=(8, 5))
        bars_pod = ax_pod.bar(pod_summary_sorted["Assign to"], pod_summary_sorted["Total_PODs"], color="orange")
        ax_pod.set_title("Total PODs per Rider")
        ax_pod.set_xlabel("Rider")
        ax_pod.set_ylabel("Total PODs")
        plt.xticks(rotation=60, ha='right')
        for bar in bars_pod:
            height = bar.get_height()
            ax_pod.annotate(f"{height}", xy=(bar.get_x() + bar.get_width() / 2, height),
                            xytext=(0, 3), textcoords="offset points", ha='center', va='bottom')
        st.pyplot(fig_pod)

        pod_img_buf = io.BytesIO()
        fig_pod.savefig(pod_img_buf, format='png', bbox_inches="tight")
        pod_img_buf.seek(0)
        st.download_button("⬇️ Download POD Chart (PNG)", pod_img_buf, "pod_chart.png", "image/png")

        # Weight chart
        pod_summary_sorted_weight = pod_summary.sort_values("Total_Weight", ascending=False)
        fig_weight, ax_weight = plt.subplots(figsize=(8, 5))
        bars_weight = ax_weight.bar(pod_summary_sorted_weight["Assign to"], pod_summary_sorted_weight["Total_Weight"], color="blue")
        ax_weight.set_title("Total Weight per Rider")
        ax_weight.set_xlabel("Rider")
        ax_weight.set_ylabel("Total Weight")
        plt.xticks(rotation=60, ha='right')
        for bar in bars_weight:
            height = bar.get_height()
            ax_weight.annotate(f"{height:.1f}", xy=(bar.get_x() + bar.get_width() / 2, height),
                               xytext=(0, 3), textcoords="offset points", ha='center', va='bottom')
        st.pyplot(fig_weight)

        weight_img_buf = io.BytesIO()
        fig_weight.savefig(weight_img_buf, format='png', bbox_inches="tight")
        weight_img_buf.seek(0)
        st.download_button("⬇️ Download Weight Chart (PNG)", weight_img_buf, "weight_chart.png", "image/png")

        output_pod = io.BytesIO()
        with pd.ExcelWriter(output_pod, engine='openpyxl') as writer:
            pod_summary.to_excel(writer, index=False, sheet_name="POD Summary")
        processed_pod = output_pod.getvalue()

        file_name_pod = f"pod_summary_{delivery_date}.xlsx"

        st.download_button("⬇️ Download POD Summary Excel", processed_pod, file_name_pod, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# -----------------------------
# Section 2: Idle Time Analysis
# -----------------------------
st.header("🕒 Idle Time, Mileage & Max Speed Analysis")

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

        summary_df_sorted_idle = summary_df.sort_values("Idle time >15 mins (hrs)", ascending=False)
        fig_idle, ax_idle = plt.subplots(figsize=(8, 5))
        bars_idle = ax_idle.bar(summary_df_sorted_idle["Rider"], summary_df_sorted_idle["Idle time >15 mins (hrs)"], color="skyblue")
        ax_idle.set_title("Idle Time >15 mins per Rider (hours)")
        ax_idle.set_xlabel("Rider")
        ax_idle.set_ylabel("Idle Time >15 mins (hrs)")
        plt.xticks(rotation=60, ha='right')
        for bar in bars_idle:
            height = bar.get_height()
            ax_idle.annotate(f"{height:.1f}", xy=(bar.get_x() + bar.get_width() / 2, height),
                             xytext=(0, 3), textcoords="offset points", ha='center', va='bottom')

        summary_df_sorted_speed = summary_df.sort_values("Max speed (km/h)", ascending=False)
        fig_speed, ax_speed = plt.subplots(figsize=(8, 5))
        bars_speed = ax_speed.bar(summary_df_sorted_speed["Rider"], summary_df_sorted_speed["Max speed (km/h)"], color="green")
        ax_speed.set_title("Max Speed per Rider (km/h)")
        ax_speed.set_xlabel("Rider")
        ax_speed.set_ylabel("Max Speed (km/h)")
        plt.xticks(rotation=60, ha='right')
        for bar in bars_speed:
            height = bar.get_height()
            ax_speed.annotate(f"{height:.0f}", xy=(bar.get_x() + bar.get_width() / 2, height),
                              xytext=(0, 3), textcoords="offset points", ha='center', va='bottom')

        summary_df_sorted_mileage = summary_df.sort_values("Total mileage (km)", ascending=False)
        fig_mileage, ax_mileage = plt.subplots(figsize=(12, 6))
        bars_mileage = ax_mileage.bar(summary_df_sorted_mileage["Rider"], summary_df_sorted_mileage["Total mileage (km)"], color="purple")
        ax_mileage.set_title("Total Mileage per Rider (km)")
        ax_mileage.set_xlabel("Rider")
        ax_mileage.set_ylabel("Total Mileage (km)")
        plt.xticks(rotation=45, ha='right', fontsize=9)
        for bar in bars_mileage:
            height = bar.get_height()
            ax_mileage.annotate(f"{height:.1f}", xy=(bar.get_x() + bar.get_width() / 2, height),
                                xytext=(0, 3), textcoords="offset points", ha='center', va='bottom')

        col1, col2 = st.columns(2)

        with col1:
            st.pyplot(fig_idle)
            idle_img_buf = io.BytesIO()
            fig_idle.savefig(idle_img_buf, format='png', bbox_inches="tight")
            idle_img_buf.seek(0)
            st.download_button("⬇️ Download Idle Time Chart (PNG)", idle_img_buf, "idle_time_chart.png", "image/png")

        with col2:
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

        output_idle = io.BytesIO()
        with pd.ExcelWriter(output_idle, engine='openpyxl') as writer:
            summary_df.to_excel(writer, index=False, sheet_name="Idle Summary")
        processed_idle = output_idle.getvalue()

        output_date = summary[0]["Date"] if summary else "unknown_date"
        file_name_idle = f"idle_time_summary_{output_date}.xlsx"

        st.download_button("⬇️ Download Idle Time Summary Excel", processed_idle, file_name_idle, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# -----------------------------
# Section 3: Cartrack Summary
# -----------------------------
st.header("🚗 Cartrack Summary")

cartrack_trip_file = st.file_uploader("Upload Summary Trip Report", type=["xlsx", "xls"], key="trip")
cartrack_fuel_file = st.file_uploader("Upload Fuel Efficiency Report", type=["xlsx", "xls"], key="fuel")

if cartrack_trip_file and cartrack_fuel_file:
    # ---------- Read Summary Trip Report ----------
    # Read all sheets as raw text to find registration
    raw_trip = pd.read_excel(cartrack_trip_file, sheet_name=0, header=None)

    # Find registration value
    registration_value = None
    for idx, row in raw_trip.iterrows():
        if str(row[0]).strip().startswith("Registration:"):
            registration_value = str(row[1]).strip()
            break

    if registration_value is None:
        st.error("❌ Could not find Registration value in trip file.")
        st.stop()

    # Read trip data with proper columns
    df_trip = pd.read_excel(cartrack_trip_file, sheet_name=0, skiprows=19)
    df_trip.columns = df_trip.columns.astype(str).str.strip()

    # Add registration column manually
    df_trip["Registration"] = registration_value

    # ---------- Read Fuel Efficiency Report ----------
    df_fuel = pd.read_excel(cartrack_fuel_file, sheet_name=0, skiprows=13)
    df_fuel.columns = df_fuel.columns.astype(str).str.strip()

    # Clean
    df_trip["Registration"] = df_trip["Registration"].astype(str).str.strip()
    df_fuel["Vehicle Registration"] = df_fuel["Vehicle Registration"].astype(str).str.strip()

    # Merge
    df_summary = pd.merge(df_trip, df_fuel, left_on="Registration", right_on="Vehicle Registration", how="left")

    # Assign drivers based on End Location
    df_summary["Assigned Driver"] = "Mohd Hairul"
    df_summary.loc[df_summary["End Location"].str.contains("Ang Mo Kio", case=False, na=False), "Assigned Driver"] = "Abdul Rahman"
    df_summary.loc[df_summary["End Location"].str.contains("Hougang|Sengkang", case=False, na=False), "Assigned Driver"] = "Abdul Rahman"

    # Vehicles that did not move
    df_summary["No Movement"] = df_summary["Trip Distance"].fillna(0) == 0

    # Show summary table
    st.subheader("📄 Cartrack Summary Table")
    st.dataframe(df_summary)

    # Fuel efficiency per driver chart
    fuel_driver_df = df_summary.groupby("Assigned Driver")["Fuel Consumed"].sum().reset_index()

    fig_fuel, ax_fuel = plt.subplots(figsize=(8, 5))
    bars = ax_fuel.bar(fuel_driver_df["Assigned Driver"], fuel_driver_df["Fuel Consumed"], color="orange")
    ax_fuel.set_title("Fuel Consumed per Driver (litres)")
    ax_fuel.set_xlabel("Driver")
    ax_fuel.set_ylabel("Fuel Consumed (litres)")
    for bar in bars:
        height = bar.get_height()
        ax_fuel.annotate(f"{height:.1f}", xy=(bar.get_x() + bar.get_width() / 2, height),
                         xytext=(0, 3), textcoords="offset points", ha='center', va='bottom')
    st.pyplot(fig_fuel)

    # Speeding incidents per driver chart
    speeding_driver_df = df_summary.groupby("Assigned Driver")["# of Events"].sum().reset_index()

    fig_speed, ax_speed = plt.subplots(figsize=(8, 5))
    bars = ax_speed.bar(speeding_driver_df["Assigned Driver"], speeding_driver_df["# of Events"], color="red")
    ax_speed.set_title("Speeding Incidents per Driver")
    ax_speed.set_xlabel("Driver")
    ax_speed.set_ylabel("Number of Events")
    for bar in bars:
        height = bar.get_height()
        ax_speed.annotate(f"{int(height)}", xy=(bar.get_x() + bar.get_width() / 2, height),
                          xytext=(0, 3), textcoords="offset points", ha='center', va='bottom')
    st.pyplot(fig_speed)

    # Highlight vehicles that did not move
    df_summary.loc[df_summary["No Movement"], "Remarks"] = "🚨 Did not move"
    df_summary.loc[~df_summary["No Movement"], "Remarks"] = ""

    # Downloadable Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_summary.to_excel(writer, index=False, sheet_name="Cartrack Summary")
    processed_data = output.getvalue()

    st.download_button(
        label="⬇️ Download Cartrack Summary Excel",
        data=processed_data,
        file_name="cartrack_summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
