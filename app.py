import streamlit as st
import pandas as pd
import io
import re
import matplotlib.pyplot as plt

st.set_page_config(page_title="Rider POD & Idle Time Analysis", layout="centered")

st.title("🚚 Rider POD & Idle Time Analysis Web App")

st.markdown("""
This tool lets you upload Excel files and get:
- ✅ Rider **POD tracking summary**
- ✅ Rider **idle time, mileage, and max speed summary**
- ✅ Idle time bar chart (converted to hours, >15 min only)

---
""")

# -----------------------------
# Section 1: POD Tracking
# -----------------------------
st.header("📦 POD Tracking Summary")

pod_file = st.file_uploader("Upload POD Excel file", type=["xlsx", "xls"], key="pod")

if pod_file:
    df_pod = pd.read_excel(pod_file)
    st.success("✅ POD file uploaded successfully!")
    st.write("Columns detected:", df_pod.columns.tolist())

    if "POD Time" in df_pod.columns and "Assign To" in df_pod.columns:
        df_pod["POD Time"] = pd.to_datetime(df_pod["POD Time"], errors='coerce')

        # Get delivery date
        if "Delivery Date" in df_pod.columns:
            delivery_date_raw = df_pod["Delivery Date"].iloc[0]
            try:
                delivery_date = pd.to_datetime(delivery_date_raw).strftime("%Y-%m-%d")
            except:
                delivery_date = "unknown_date"
        else:
            delivery_date = "unknown_date"

        # Group and summarize
        pod_summary = df_pod.groupby("Assign To").agg(
            Earliest_POD=("POD Time", "min"),
            Latest_POD=("POD Time", "max"),
            Total_PODs=("POD Time", "count")
        ).reset_index()

        st.subheader("📄 POD Summary Table")
        st.dataframe(pod_summary)

        # Convert to Excel
        output_pod = io.BytesIO()
        with pd.ExcelWriter(output_pod, engine='openpyxl') as writer:
            pod_summary.to_excel(writer, index=False, sheet_name="POD Summary")
        processed_pod = output_pod.getvalue()

        file_name_pod = f"pod_summary_{delivery_date}.xlsx"

        st.download_button(
            label="⬇️ Download POD Summary Excel",
            data=processed_pod,
            file_name=file_name_pod,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("❌ Required columns 'Assign To' and 'POD Time' not found in this file.")

# -----------------------------
# Section 2: Idle Time Analysis
# -----------------------------
st.header("🕒 Idle Time & Mileage Analysis")

rider_files = st.file_uploader("Upload multiple rider Excel files", type=["xlsx", "xls"], accept_multiple_files=True, key="idle")

if rider_files:
    summary = []

    for file in rider_files:
        # Extract date from filename
        date_match = re.search(r'\d{4}-\d{2}-\d{2}', file.name)
        if date_match:
            date_str = date_match.group(0)
        else:
            date_str = "unknown_date"

        xl = pd.ExcelFile(file)
        rider_name = xl.sheet_names[0]
        df = pd.read_excel(file, sheet_name=rider_name)

        if "Time" not in df.columns or "Mileage (km)" not in df.columns or "Speed (km/h)" not in df.columns:
            st.error(f"❌ File {file.name} is missing required columns.")
            continue

        df['Time'] = pd.to_datetime(df['Time'], format='%I:%M:%S %p', errors='coerce')
        df['Idle'] = df['Mileage (km)'] == 0

        idle_periods = []
        current_start = None

        for idx, row in df.iterrows():
            if row['Idle']:
                if current_start is None:
                    current_start = row['Time']
            else:
                if current_start is not None:
                    idle_periods.append((current_start, row['Time']))
                    current_start = None
        if current_start is not None:
            idle_periods.append((current_start, df['Time'].iloc[-1]))

        idle_durations = [(end - start).total_seconds() / 60 for start, end in idle_periods]
        total_idle = sum(idle_durations)
        over_15 = [d for d in idle_durations if d > 15]
        total_over_15 = sum(over_15)
        num_over_15 = len(over_15)
        total_mileage = df['Mileage (km)'].sum()
        max_speed = df['Speed (km/h)'].max()

        summary.append({
            "File": file.name,
            "Rider": rider_name,
            "Date": date_str,
            "Total idle time (mins)": total_idle,
            "Idle time >15 mins (mins)": total_over_15,
            "Num idle periods >15 mins": num_over_15,
            "Total mileage (km)": total_mileage,
            "Max speed (km/h)": max_speed
        })

    if summary:
        summary_df = pd.DataFrame(summary)

        # Create formatted "X hr Y min" column for idle >15 mins
        def format_hours_mins(x):
            if x == 0 or pd.isna(x):
                return "0 hr 0 min"
            hours = int(x // 60)
            mins = int(x % 60)
            return f"{hours} hr {mins} min"

        summary_df["Idle >15 mins (formatted)"] = summary_df["Idle time >15 mins (mins)"].apply(format_hours_mins)

        # Convert idle >15 mins (mins) to hours for chart
        summary_df["Idle time >15 mins (hrs)"] = summary_df["Idle time >15 mins (mins)"] / 60

        # Sort for chart
        summary_df_sorted = summary_df.sort_values("Idle time >15 mins (hrs)", ascending=False)

        # Create bar chart
        fig, ax = plt.subplots(figsize=(8, 5))
        ax.bar(summary_df_sorted["Rider"], summary_df_sorted["Idle time >15 mins (hrs)"], color="skyblue")
        ax.set_title("Idle Time >15 mins per Rider (hours)")
        ax.set_xlabel("Rider")
        ax.set_ylabel("Idle Time >15 mins (hrs)")
        plt.xticks(rotation=45)

        st.pyplot(fig)

        # Remove numeric hours column before display & export
        summary_df = summary_df.drop(columns=["Idle time >15 mins (hrs)"])

        st.subheader("📄 Idle Time Summary Table")
        st.dataframe(summary_df)

        # Convert to Excel
        output_idle = io.BytesIO()
        with pd.ExcelWriter(output_idle, engine='openpyxl') as writer:
            summary_df.to_excel(writer, index=False, sheet_name="Idle Summary")
        processed_idle = output_idle.getvalue()

        # Use date from first file for filename
        output_date = summary[0]["Date"] if summary else "unknown_date"
        file_name_idle = f"idle_time_summary_{output_date}.xlsx"

        st.download_button(
            label="⬇️ Download Idle Time Summary Excel",
            data=processed_idle,
            file_name=file_name_idle,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
