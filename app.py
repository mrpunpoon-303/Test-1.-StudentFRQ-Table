import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

def load_data(file):
    try:
        data = pd.read_excel(file)
        data["Start_Date_time"] = pd.to_datetime(data["Start_Date_time"], errors="coerce")
        return data
    except Exception as e:
        st.error(f"Error loading file: {e}")
        return None

def create_frequency_table(data, period=None, start_period=None, end_period=None, max_upper=10):
    if period:
        data_filtered = data[data["Start_Date_time"].dt.to_period("M").astype(str) == period]
    elif start_period and end_period:
        data_filtered = data[
            (data["Start_Date_time"].dt.to_period("M").astype(str) >= start_period) &
            (data["Start_Date_time"].dt.to_period("M").astype(str) <= end_period)
        ]
    else:
        return None

    data_filtered = data_filtered[~data_filtered["Class_Name"].str.contains("Self Practice", case=False, na=False)]
    booking_frequencies = data_filtered.groupby("Id_Person").size()

    table = pd.DataFrame({
        "Freq": list(range(1, max_upper + 1)) + [f">{max_upper}"],
        "#Students": [sum(booking_frequencies == i) for i in range(1, max_upper + 1)] + 
                    [sum(booking_frequencies > max_upper)],
        "Cum 1->": [sum(booking_frequencies <= i) for i in range(1, max_upper + 1)] + 
                   [len(booking_frequencies)],
        "Cum ->End": [len(booking_frequencies) - sum(booking_frequencies <= i) for i in range(1, max_upper + 1)] + 
                     [sum(booking_frequencies > max_upper)]
    })

    def get_student_details(freq):
        if isinstance(freq, str) and freq.startswith(">"):
            ids = booking_frequencies[booking_frequencies > max_upper].index
        else:
            ids = booking_frequencies[booking_frequencies == freq].index
        names = data_filtered[data_filtered["Id_Person"].isin(ids)]["FirstName"].drop_duplicates()
        return ", ".join([f"{name} : {id}" for name, id in zip(names, ids)])

    table["Details"] = [get_student_details(freq) for freq in table["Freq"]]
    return table

def plot_histogram(table):
    fig, ax = plt.subplots(figsize=(12, 8))
    bars = ax.bar(table["Freq"].astype(str), table["#Students"], color='skyblue', edgecolor='black', width=0.7)

    student_counts = table.loc[table["Freq"].apply(lambda x: str(x).isdigit()), 
                             ["Freq", "#Students"]].copy()
    student_counts["Freq"] = student_counts["Freq"].astype(int)
    expanded = student_counts.loc[student_counts.index.repeat(student_counts["#Students"])]
    mean_val = expanded["Freq"].mean()
    median_val = expanded["Freq"].median()

    ax.axvline(mean_val, color='red', linestyle='--', linewidth=1, label=f'Mean: {mean_val:.2f}')
    ax.axvline(median_val, color='green', linestyle='--', linewidth=1, label=f'Median: {median_val:.2f}')

    for bar, count in zip(bars, table["#Students"]):
        height = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2, height + 0.5, 
               str(count), ha='center', fontsize=10)

    ax.set_xlabel("Frequency of Bookings", fontsize=12)
    ax.set_ylabel("Number of Students", fontsize=12)
    ax.set_xticklabels(table["Freq"].astype(str), fontsize=10, rotation=45)
    ax.grid(axis='y', linestyle='--', alpha=0.7)
    plt.legend()
    plt.tight_layout()
    
    return fig

def main():
    st.set_page_config(layout="wide")
    st.title("Booking Frequency Analysis")
    
    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
    if uploaded_file:
        data = load_data(uploaded_file)
        if data is not None:
            col1, col2, col3 = st.columns([1,1,1])
            
            with col1:
                analysis_type = st.radio("Analysis Type:", ["Monthly", "Range"])
            with col2:
                max_upper = st.number_input("Max Upper Bound:", min_value=1, value=15)
            
            periods = sorted(data["Start_Date_time"].dt.to_period("M").astype(str).unique())
            
            if analysis_type == "Monthly":
                period = st.selectbox("Select Period:", periods)
                table = create_frequency_table(data, period=period, max_upper=max_upper)
                title = f"Booking Frequency Report for {period}"
            else:
                col1, col2 = st.columns(2)
                with col1:
                    start_period = st.selectbox("Start Period:", periods)
                with col2:
                    end_period = st.selectbox("End Period:", periods)
                table = create_frequency_table(data, start_period=start_period, 
                                            end_period=end_period, max_upper=max_upper)
                title = f"Booking Frequency Report from {start_period} to {end_period}"
            
            if table is not None:
                st.subheader(title)
                st.pyplot(plot_histogram(table))
                st.dataframe(
                    table,
                    use_container_width=True,
                    column_config={
                        "Details": st.column_config.TextColumn(
                            "Details",
                            width="large",
                            help="Student details"
                        )
                    }
                )

if __name__ == "__main__":
    main()
