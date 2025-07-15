import streamlit as st
import pandas as pd
import io
import re
import matplotlib.pyplot as plt
import datetime
import difflib

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
    df_pod.columns = df_pod.columns.str.strip()

    st.success("✅ POD file uploaded successfully!")

    # check required columns
    if all(col in df_pod.columns for col in ["Assign to","POD Time","Weight","Delivery Date"]):
        # parse
        df_pod["Delivery Date"] = pd.to_datetime(df_pod["Delivery Date"], errors="coerce")
        df_pod["POD Time"]      = pd.to_datetime(df_pod["POD Time"], errors="coerce").dt.time
        df_pod["POD DateTime"] = df_pod.apply(
            lambda r: datetime.datetime.combine(r["Delivery Date"], r["POD Time"]) 
                      if pd.notnull(r["Delivery Date"]) and pd.notnull(r["POD Time"]) 
                      else pd.NaT,
            axis=1
        )
        # filename date
        try:
            delivery_date = df_pod["Delivery Date"].mode()[0].strftime("%Y-%m-%d")
        except:
            delivery_date = "unknown_date"

        # summary
        pod_summary = df_pod.groupby("Assign to").agg(
            Earliest_POD=("POD DateTime","min"),
            Latest_POD=("POD DateTime","max"),
            Total_PODs=("POD DateTime","count"),
            Total_Weight=("Weight","sum")
        ).reset_index()

        # display table
        st.subheader("📄 POD Summary Table")
        st.dataframe(pod_summary)

        # — chart 1: POD count
        pod_summary_sorted = pod_summary.sort_values("Total_PODs", ascending=False)
        fig_pod, ax_pod = plt.subplots(figsize=(8,5))
        bars = ax_pod.bar(pod_summary_sorted["Assign to"], pod_summary_sorted["Total_PODs"], color="orange")
        ax_pod.set_title("Total PODs per Rider")
        ax_pod.set_xlabel("Rider")
        ax_pod.set_ylabel("Total PODs")
        plt.xticks(rotation=60, ha="right")
        for b in bars:
            ax_pod.annotate(f"{b.get_height():.0f}",
                            (b.get_x()+b.get_width()/2, b.get_height()),
                            ha="center", va="bottom", xytext=(0,3), textcoords="offset points")
        st.pyplot(fig_pod)

        # save chart to buffer
        pod_img_buf = io.BytesIO()
        fig_pod.savefig(pod_img_buf, format="png", bbox_inches="tight")
        pod_img_buf.seek(0)
        st.download_button("⬇️ Download POD Chart (PNG)", pod_img_buf, "pod_chart.png", "image/png")

        # — chart 2: total weight
        pod_summary_sorted_w = pod_summary.sort_values("Total_Weight", ascending=False)
        fig_wt, ax_wt = plt.subplots(figsize=(8,5))
        bars = ax_wt.bar(pod_summary_sorted_w["Assign to"], pod_summary_sorted_w["Total_Weight"], color="blue")
        ax_wt.set_title("Total Weight per Rider")
        ax_wt.set_xlabel("Rider")
        ax_wt.set_ylabel("Total Weight")
        plt.xticks(rotation=60, ha="right")
        for b in bars:
            ax_wt.annotate(f"{b.get_height():.1f}",
                           (b.get_x()+b.get_width()/2, b.get_height()),
                           ha="center", va="bottom", xytext=(0,3), textcoords="offset points")
        st.pyplot(fig_wt)

        weight_img_buf = io.BytesIO()
        fig_wt.savefig(weight_img_buf, format="png", bbox_inches="tight")
        weight_img_buf.seek(0)
        st.download_button("⬇️ Download Weight Chart (PNG)", weight_img_buf, "weight_chart.png", "image/png")

        # — embed both charts into Excel
        excel_buf = io.BytesIO()
        with pd.ExcelWriter(excel_buf, engine="xlsxwriter") as writer:
            pod_summary.to_excel(writer, sheet_name="POD Summary", index=False)
            wb  = writer.book
            ws  = writer.sheets["POD Summary"]
            # adjust these anchors as you like
            ws.insert_image("H2",  "pod_chart.png",    {"image_data": pod_img_buf})
            ws.insert_image("H20", "weight_chart.png", {"image_data": weight_img_buf})
        excel_buf.seek(0)
        st.download_button(
            "⬇️ Download POD Summary Excel",
            data=excel_buf,
            file_name=f"pod_summary_{delivery_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("❌ Required columns missing in your POD file.")


# -----------------------------
# Section 2: Idle Time Analysis
# -----------------------------
st.header("🕒 Idle Time, Mileage & Max Speed Analysis")
rider_files = st.file_uploader(
    "Upload multiple rider Excel files (vehicle route)",
    type=["xlsx","xls"],
    accept_multiple_files=True,
    key="idle"
)

if rider_files:
    summary = []
    for f in rider_files:
        # infer date from filename
        m = re.search(r"\d{4}-\d{2}-\d{2}", f.name)
        date_str = m.group(0) if m else "unknown_date"

        xl = pd.ExcelFile(f)
        rider_name = xl.sheet_names[0]
        df = pd.read_excel(f, sheet_name=rider_name)

        if not all(c in df.columns for c in ["Time","Mileage (km)","Speed (km/h)"]):
            st.error(f"❌ File {f.name} missing required columns.")
            continue

        # parse and flag idle
        df["Time"]    = pd.to_datetime(df["Time"], format="%I:%M:%S %p", errors="coerce")
        df["Idle"]    = df["Mileage (km)"] == 0
        work_start    = datetime.time(8,30)
        work_end      = datetime.time(17,30)
        df["t_only"]  = df["Time"].dt.time
        df_working    = df[(df["t_only"]>=work_start)&(df["t_only"]<=work_end)]
        total_mileage = df_working["Mileage (km)"].sum()

        if total_mileage == 0:
            total_idle = total_over = count_over = max_speed = 0
            status = "Not working for the day"
        else:
            # build idle periods
            periods = []
            start = None
            for _,row in df.iterrows():
                t = row["Time"].time()
                if t<work_start or t>work_end:
                    if start:
                        periods.append((start,row["Time"]))
                        start=None
                    continue
                if row["Idle"]:
                    if not start:
                        start = row["Time"]
                else:
                    if start:
                        periods.append((start,row["Time"]))
                        start=None
            if start:
                periods.append((start, df["Time"].iloc[-1]))

            durs = [(end-start).total_seconds()/60 for start,end in periods]
            total_idle = sum(durs)
            over15     = [d for d in durs if d>15]
            total_over = sum(over15)
            count_over = len(over15)
            max_speed  = df_working["Speed (km/h)"].max()
            status     = "Working for the day"

        summary.append({
            "File": f.name,
            "Rider": rider_name,
            "Date": date_str,
            "Total idle time (mins)": total_idle,
            "Idle time >15 mins (mins)": total_over,
            "Num idle periods >15 mins": count_over,
            "Total mileage (km)": total_mileage,
            "Max speed (km/h)": max_speed,
            "Status": status
        })

    if summary:
        df_sum = pd.DataFrame(summary)

        # format and sort
        df_sum["Idle >15 mins (hrs)"] = df_sum["Idle time >15 mins (mins)"]/60
        idle_sorted  = df_sum.sort_values("Idle >15 mins (hrs)", ascending=False)
        speed_sorted = df_sum.sort_values("Max speed (km/h)", ascending=False)
        mil_sorted   = df_sum.sort_values("Total mileage (km)", ascending=False)

        # — chart: Idle >15 hrs
        fig_idle, ax_idle = plt.subplots(figsize=(8,5))
        bars = ax_idle.bar(idle_sorted["Rider"], idle_sorted["Idle >15 mins (hrs)"], color="skyblue")
        ax_idle.set_title("Idle Time >15 mins per Rider (hrs)")
        ax_idle.set_xlabel("Rider"); ax_idle.set_ylabel("Hours")
        plt.xticks(rotation=60, ha="right")
        for b in bars:
            ax_idle.annotate(f"{b.get_height():.1f}", 
                             (b.get_x()+b.get_width()/2, b.get_height()),
                             ha="center", va="bottom", xytext=(0,3), textcoords="offset points")

        # — chart: Max speed
        fig_speed, ax_speed = plt.subplots(figsize=(8,5))
        bars = ax_speed.bar(speed_sorted["Rider"], speed_sorted["Max speed (km/h)"], color="green")
        ax_speed.set_title("Max Speed per Rider (km/h)")
        ax_speed.set_xlabel("Rider"); ax_speed.set_ylabel("km/h")
        plt.xticks(rotation=60, ha="right")
        for b in bars:
            ax_speed.annotate(f"{b.get_height():.0f}",
                              (b.get_x()+b.get_width()/2, b.get_height()),
                              ha="center", va="bottom", xytext=(0,3), textcoords="offset points")

        # — chart: Mileage
        fig_mil, ax_mil = plt.subplots(figsize=(12,6))
        bars = ax_mil.bar(mil_sorted["Rider"], mil_sorted["Total mileage (km)"], color="purple")
        ax_mil.set_title("Total Mileage per Rider (km)")
        ax_mil.set_xlabel("Rider"); ax_mil.set_ylabel("km")
        plt.xticks(rotation=45, ha="right", fontsize=9)
        for b in bars:
            ax_mil.annotate(f"{b.get_height():.1f}",
                            (b.get_x()+b.get_width()/2, b.get_height()),
                            ha="center", va="bottom", xytext=(0,3), textcoords="offset points")

        # display & download each
        col1, col2 = st.columns(2)
        # Idle
        with col1:
            st.pyplot(fig_idle)
            idle_buf = io.BytesIO(); fig_idle.savefig(idle_buf, format="png", bbox_inches="tight"); idle_buf.seek(0)
            st.download_button("⬇️ Download Idle Chart (PNG)", idle_buf, "idle_chart.png", "image/png")
        # Speed
        with col2:
            st.pyplot(fig_speed)
            speed_buf = io.BytesIO(); fig_speed.savefig(speed_buf, format="png", bbox_inches="tight"); speed_buf.seek(0)
            st.download_button("⬇️ Download Speed Chart (PNG)", speed_buf, "speed_chart.png", "image/png")
        # Mileage below
        st.pyplot(fig_mil)
        mil_buf = io.BytesIO(); fig_mil.savefig(mil_buf, format="png", bbox_inches="tight"); mil_buf.seek(0)
        st.download_button("⬇️ Download Mileage Chart (PNG)", mil_buf, "mileage_chart.png", "image/png")

        # final table
        st.subheader("📄 Idle Time Summary Table")
        st.dataframe(df_sum.drop(columns=["Idle >15 mins (hrs)"]))

        # embed into Excel
        excel_idle = io.BytesIO()
        with pd.ExcelWriter(excel_idle, engine="xlsxwriter") as writer:
            df_sum.to_excel(writer, sheet_name="Idle Summary", index=False)
            wb  = writer.book
            ws  = writer.sheets["Idle Summary"]
            ws.insert_image("J2",  "idle_chart.png",   {"image_data": idle_buf})
            ws.insert_image("J20", "speed_chart.png",  {"image_data": speed_buf})
            ws.insert_image("J38", "mileage_chart.png",{"image_data": mil_buf})
        excel_idle.seek(0)
        st.download_button(
            "⬇️ Download Idle Time Summary Excel",
            data=excel_idle,
            file_name=f"idle_time_summary_{df_sum.iloc[0]['Date']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
