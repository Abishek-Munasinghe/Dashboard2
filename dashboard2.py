import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import seaborn as sns


# Paths to both Excel files
current_progress_file = r"Gather missing application escalation information Tracker - FY25 (2).xlsx"
previous_progress_file = r"Gather missing application escalation information Tracker - FY24 - Copy.xlsx"


def load_and_normalize(file_path):
    """
    Load the entire Excel file and normalize its column names.
    """
    try:
        # Load Excel file data
        df = pd.read_excel(file_path)
        
        # Normalize column names
        df.columns = df.columns.str.strip().str.title()
        
        return df
    except Exception as e:
        st.error(f"Error reading {file_path}: {e}")
        return None


def main():
    st.title("Escalation contacts Gathering goal Progress Dashboard")
    st.write("Visualizing and comparing sprint progress before and during the sprint.")

    # Load and normalize data from both files
    st.write("Loading data from both progress reports...")
    current_df = load_and_normalize(current_progress_file)
    previous_df = load_and_normalize(previous_progress_file)

    if current_df is None or previous_df is None:
        st.error("Error loading data from files. Please check file paths and try again.")
        return

    # Ensure necessary columns exist
    if "Assign To" not in current_df.columns or "Status" not in current_df.columns:
        st.error("The required columns are not found in the current progress data.")
        return

    if "Assign To" not in previous_df.columns or "Status" not in previous_df.columns:
        st.error("The required columns are not found in the previous progress data.")
        return

    # Group and count statuses for each group
    current_summary = current_df.groupby(["Assign To", "Status"]).size().reset_index(name="Count")
    previous_summary = previous_df.groupby(["Assign To", "Status"]).size().reset_index(name="Count")

    # Plot Graphs Separately
    st.write("### Current Sprint Progress Graph")
    plt.figure(figsize=(12, 6))
    sns.barplot(data=current_summary, x="Assign To", y="Count", hue="Status")
    plt.xticks(rotation=45)
    plt.title("Current Sprint Progress by Assigned Person & Status")
    plt.ylabel("Task Count")
    plt.xlabel("Assigned Person")
    st.pyplot(plt)

    st.write("### Previous Sprint Progress Graph")
    plt.figure(figsize=(12, 6))
    sns.barplot(data=previous_summary, x="Assign To", y="Count", hue="Status")
    plt.xticks(rotation=45)
    plt.title("Previous Sprint Progress by Assigned Person & Status")
    plt.ylabel("Task Count")
    plt.xlabel("Assigned Person")
    st.pyplot(plt)

    # Comparison Section
    st.write("### Comparison of Current vs Previous Status Counts")
    # Merge the two summaries for comparison
    comparison_df = pd.merge(
        current_summary,
        previous_summary,
        on=["Assign To", "Status"],
        how="outer",
        suffixes=("_Current", "_Previous"),
    ).fillna(0)

    # Prepare melted data for visualization
    comparison_df_melted = comparison_df.melt(
        id_vars=["Assign To", "Status"], 
        value_vars=["Count_Current", "Count_Previous"], 
        var_name="Period", 
        value_name="Count"
    )

    # Plot the comparison graph
    plt.figure(figsize=(12, 6))
    sns.barplot(data=comparison_df_melted, x="Assign To", y="Count", hue="Period", hue_order=["Count_Previous", "Count_Current"])
    plt.xticks(rotation=45)
    plt.title("Comparison of Progress Before & During Sprint")
    plt.ylabel("Task Count")
    plt.xlabel("Assigned Person")
    st.pyplot(plt)

    # Display the summary stats
    st.write("### Summary Statistics")
    st.write(comparison_df)


if __name__ == "__main__":
    main()
