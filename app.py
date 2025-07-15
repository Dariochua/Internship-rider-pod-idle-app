import streamlit as st
import pandas as pd
import io
import re
import matplotlib.pyplot as plt
import datetime
import xlsxwriter

st.set_page_config(page_title="Rider POD & Idle Time Analysis", layout="centered")

st.title("🚚 Rider POD & Idle Time Analysis Web App")
st.markdown("""
This tool lets you upload Detrack Excel files and get:
- Rider **POD tracking summary + charts** (displayed in-app)
- Rider **idle time, mileage, and max speed summary + charts** (displayed in-app)
- Downloadable Excel reports with **embedded charts**, complete with data labels and custom colors
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

    if {"Assign to","POD Time","Weight","Delivery Date"}.issubset(df_pod.columns):
        df_pod["Delivery Date"] = pd.to_datetime(df_pod["Delivery Date"], errors='coerce')
        df_pod["POD Time"]      = pd.to_datetime(df_pod["POD Time"], errors='coerce').dt.time
        df_pod["POD DateTime"]  = df_pod.apply(
            lambda r: datetime.datetime.combine(r["Delivery Date"], r["POD Time"])
                      if pd.notnull(r["Delivery Date"]) and pd.notnull(r["POD Time"]) else pd.NaT,
            axis=1
        )
        mode_date = df_pod["Delivery Date"].mode()[0]
        delivery_date = mode_date.strftime("%Y-%m-%d")

        pod_summary = df_pod.groupby("Assign to").agg(
            Earliest_POD  = ("POD DateTime","min"),
            Latest_POD    = ("POD DateTime","max"),
            Total_PODs    = ("POD DateTime","count"),
            Total_Weight  = ("Weight","sum")
        ).reset_index()

        # In-app plots
        fig1, ax1 = plt.subplots(figsize=(8,5))
        bars1 = ax1.bar(pod_summary["Assign to"], pod_summary["Total_PODs"], color="#FFA500")
        ax1.set_title("Total PODs per Rider", name_font={'size':14,'bold':True})
        ax1.set_xlabel("Rider", name_font={'size':12})
        ax1.set_ylabel("POD Count", name_font={'size':12})
        plt.xticks(rotation=60, ha='right')
        for b in bars1:
            ax1.annotate(int(b.get_height()), xy=(b.get_x()+b.get_width()/2, b.get_height()),
                         xytext=(0,3), textcoords="offset points", ha='center', va='bottom')
        st.pyplot(fig1)

        fig2, ax2 = plt.subplots(figsize=(8,5))
        bars2 = ax2.bar(pod_summary["Assign to"], pod_summary["Total_Weight"], color="#0000FF")
        ax2.set_title("Total Weight per Rider", name_font={'size':14,'bold':True})
        ax2.set_xlabel("Rider", name_font={'size':12})
        ax2.set_ylabel("Weight", name_font={'size':12})
        plt.xticks(rotation=60, ha='right')
        for b in bars2:
            ax2.annotate(f"{b.get_height():.1f}", xy=(b.get_x()+b.get_width()/2, b.get_height()),
                         xytext=(0,3), textcoords="offset points", ha='center', va='bottom')
        st.pyplot(fig2)

        # Excel with embedded charts
        output_pod = io.BytesIO()
        with pd.ExcelWriter(output_pod, engine='xlsxwriter') as writer:
            pod_summary.to_excel(writer, index=False, sheet_name="POD Summary")
            wb  = writer.book
            ws  = writer.sheets["POD Summary"]
            n   = len(pod_summary)

            # Chart: Total PODs
            c1 = wb.add_chart({'type':'column'})
            c1.add_series({
                'name':       'Total PODs',
                'categories': ['POD Summary', 1, 0, n, 0],
                'values':     ['POD Summary', 1, 2, n, 2],
                'fill':       {'color':'#FFA500'},
                'data_labels':{'value':True}
            })
            c1.set_title({'name':'Total PODs per Rider','name_font':{'size':14,'bold':True}})
            c1.set_x_axis({'name':'Rider','name_font':{'size':12}})
            c1.set_y_axis({'name':'POD Count','name_font':{'size':12}})
            ws.insert_chart('F2', c1, {'x_scale':1.5,'y_scale':1.5})

            # Chart: Total Weight
            c2 = wb.add_chart({'type':'column'})
            c2.add_series({
                'name':       'Total Weight',
                'categories': ['POD Summary', 1, 0, n, 0],
                'values':     ['POD Summary', 1, 3, n, 3],
                'fill':       {'color':'#0000FF'},
                'data_labels':{'value':True}
            })
            c2.set_title({'name':'Total Weight per Rider','name_font':{'size':14,'bold':True}})
            c2.set_x_axis({'name':'Rider','name_font':{'size':12}})
            c2.set_y_axis({'name':'Weight','name_font':{'size':12}})
            ws.insert_chart('F20', c2, {'x_scale':1.5,'y_scale':1.5})

        st.download_button(
            "⬇️ Download POD Summary with Charts (Excel)",
            output_pod.getvalue(),
            f"pod_summary_{delivery_date}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.error("❌ Required columns not found.")

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
    records = []
    for file in rider_files:
        date_str = (re.search(r'\d{4}-\d{2}-\d{2}', file.name) or ["unknown_date"])[0]
        xl = pd.ExcelFile(file)
        rider = xl.sheet_names[0]
        df = pd.read_excel(file, sheet_name=rider)

        if not {"Time","Mileage (km)","Speed (km/h)"}.issubset(df.columns):
            st.error(f"❌ {file.name} missing columns.")
            continue

        df['Time'] = pd.to_datetime(df['Time'], format='%I:%M:%S %p', errors='coerce')
        df['Idle'] = df['Mileage (km)']==0

        start, end = datetime.time(8,30), datetime.time(17,30)
        df['t_only'] = df['Time'].dt.time
        working = df[(df['t_only']>=start)&(df['t_only']<=end)]

        total_mileage = working['Mileage (km)'].sum()
        if total_mileage==0:
            total_idle = over15 = cnt15 = max_spd = 0
            status = "Not working"
        else:
            periods = []
            cur_start=None
            for _,r in df.iterrows():
                t = r['Time'].time()
                if t<start or t>end:
                    if cur_start:
                        periods.append((cur_start,r['Time']))
                        cur_start=None
                    continue
                if r['Idle']:
                    if not cur_start: cur_start=r['Time']
                else:
                    if cur_start:
                        periods.append((cur_start,r['Time']))
                        cur_start=None
            if cur_start: periods.append((cur_start,df['Time'].iloc[-1]))

            dur = [(e-s).total_seconds()/60 for s,e in periods]
            total_idle = sum(dur)
            over15 = [d for d in dur if d>15]
            over15_sum = sum(over15)
            cnt15 = len(over15)
            max_spd = working['Speed (km/h)'].max()
            status = "Working"

        records.append({
            "File": file.name,
            "Rider": rider,
            "Date": date_str,
            "Total idle time (mins)": total_idle,
            "Idle time >15 mins (mins)": sum(over15) if total_mileage else 0,
            "Num idle periods >15 mins": cnt15,
            "Total mileage (km)": total_mileage,
            "Max speed (km/h)": max_spd,
            "Status": status
        })

    if records:
        df_sum = pd.DataFrame(records)
        df_sum["Idle time >15 mins (hrs)"] = df_sum["Idle time >15 mins (mins)"] / 60

        # In-app plots
        sorted_idle   = df_sum.sort_values("Idle time >15 mins (hrs)", ascending=False)
        sorted_speed  = df_sum.sort_values("Max speed (km/h)", ascending=False)
        sorted_miles  = df_sum.sort_values("Total mileage (km)", ascending=False)

        fig3, ax3 = plt.subplots(figsize=(8,5))
        b3 = ax3.bar(sorted_idle["Rider"], sorted_idle["Idle time >15 mins (hrs)"], color="#87CEEB")
        ax3.set_title("Idle >15 mins per Rider (hrs)", name_font={'size':14,'bold':True})
        ax3.set_xlabel("Rider", name_font={'size':12})
        ax3.set_ylabel("Idle (hrs)", name_font={'size':12})
        plt.xticks(rotation=60, ha='right')
        for b in b3:
            ax3.annotate(f"{b.get_height():.1f}", xy=(b.get_x()+b.get_width()/2,b.get_height()),
                         xytext=(0,3), textcoords="offset points", ha='center', va='bottom')
        st.pyplot(fig3)

        fig4, ax4 = plt.subplots(figsize=(8,5))
        b4 = ax4.bar(sorted_speed["Rider"], sorted_speed["Max speed (km/h)"], color="#008000")
        ax4.set_title("Max Speed per Rider (km/h)", name_font={'size':14,'bold':True})
        ax4.set_xlabel("Rider", name_font={'size':12})
        ax4.set_ylabel("Max Speed", name_font={'size':12})
        plt.xticks(rotation=60, ha='right')
        for b in b4:
            ax4.annotate(int(b.get_height()), xy=(b.get_x()+b.get_width()/2,b.get_height()),
                         xytext=(0,3), textcoords="offset points", ha='center', va='bottom')
        st.pyplot(fig4)

        fig5, ax5 = plt.subplots(figsize=(12,6))
        b5 = ax5.bar(sorted_miles["Rider"], sorted_miles["Total mileage (km)"], color="#800080")
        ax5.set_title("Total Mileage per Rider (km)", name_font={'size':14,'bold':True})
        ax5.set_xlabel("Rider", name_font={'size':12})
        ax5.set_ylabel("Mileage (km)", name_font={'size':12})
        plt.xticks(rotation=45, ha='right', fontsize=9)
        for b in b5:
            ax5.annotate(f"{b.get_height():.1f}", xy=(b.get_x()+b.get_width()/2,b.get_height()),
                         xytext=(0,3), textcoords="offset points", ha='center', va='bottom')
        st.pyplot(fig5)

        # Excel export with embedded charts
        out_idle = io.BytesIO()
        with pd.ExcelWriter(out_idle, engine='xlsxwriter') as writer:
            df_sum.to_excel(writer, index=False, sheet_name="Idle Summary")
            wb = writer.book
            ws = writer.sheets["Idle Summary"]
            n  = len(df_sum)

            # Idle chart
            ch1 = wb.add_chart({'type':'column'})
            ch1.add_series({
                'name':        'Idle >15 mins (hrs)',
                'categories':  ['Idle Summary', 1, 1, n, 1],
                'values':      ['Idle Summary', 1, df_sum.columns.get_loc("Idle time >15 mins (hrs)"), n, df_sum.columns.get_loc("Idle time >15 mins (hrs)")],
                'fill':        {'color':'#87CEEB'},
                'data_labels': {'value':True}
            })
            ch1.set_title({'name':'Idle >15 mins per Rider (hrs)','name_font':{'size':14,'bold':True}})
            ws.insert_chart('K2', ch1, {'x_scale':1.3,'y_scale':1.3})

            # Speed chart
            ch2 = wb.add_chart({'type':'column'})
            ch2.add_series({
                'name':        'Max Speed (km/h)',
                'categories':  ['Idle Summary', 1, 1, n, 1],
                'values':      ['Idle Summary', 1, df_sum.columns.get_loc("Max speed (km/h)"), n, df_sum.columns.get_loc("Max speed (km/h)")],
                'fill':        {'color':'#008000'},
                'data_labels': {'value':True}
            })
            ch2.set_title({'name':'Max Speed per Rider','name_font':{'size':14,'bold':True}})
            ws.insert_chart('K20', ch2, {'x_scale':1.3,'y_scale':1.3})

            # Mileage chart
            ch3 = wb.add_chart({'type':'column'})
            ch3.add_series({
                'name':        'Total Mileage (km)',
                'categories':  ['Idle Summary', 1, 1, n, 1],
                'values':      ['Idle Summary', 1, df_sum.columns.get_loc("Total mileage (km)"), n, df_sum.columns.get_loc("Total mileage (km)")],
                'fill':        {'color':'#800080'},
                'data_labels': {'value':True}
            })
            ch3.set_title({'name':'Total Mileage per Rider (km)','name_font':{'size':14,'bold':True}})
            ws.insert_chart('K38', ch3, {'x_scale':1.3,'y_scale':1.3})

        st.download_button(
            "⬇️ Download Idle Summary with Charts (Excel)",
            out_idle.getvalue(),
            f"idle_time_summary_{df_sum.loc[0,'Date']}.xlsx",
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
