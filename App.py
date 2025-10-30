import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io
import zipfile
from datetime import datetime
import numpy as np

# Model data
RD93 = {
    'x': [0,0.5,1,1.5,2,2.5,3,3.1,3.2,3.3,3.4,3.5,3.6,3.7,3.8,3.9,4,4.1,4.2,4.3,4.4,4.5,4.6,4.7,4.8,4.9,5,5.1,5.2,5.3,5.4,5.5,5.6,5.7,5.8,5.9,6,6.1,6.2,6.3,6.4,6.5,6.6,6.7,6.8,6.9,7,7.1,7.2,7.3,7.4,7.5,7.6,7.7,7.8,7.9,8,8.1,8.2,8.3,8.4,8.5,8.6,8.7,8.8,8.9,9,9.1,9.2,9.3,9.4,9.5,9.6,9.7,9.8,9.9,10,10.1,10.2,10.3,10.4,10.5,10.6,10.7,10.8,10.9,11,11.1,11.2,11.3,11.4,11.5,11.6,11.7,11.8,11.9,12,12.1,12.2,12.3,12.4,12.5,12.6,12.7,12.8,12.9,13,13.5,14,14.5,15,15.5,16,16.5,17.1],
    'y': [0,0,0,0,0,0,1,8,14,21,27,34,42,51,59,68,77,89,102,115,128,141,154,168,182,196,209,222,234,247,260,272,292,312,332,353,373,397,422,446,471,496,523,550,577,604,631,660,689,718,747,776,812,848,883,919,954,991,1027,1064,1101,1137,1178,1218,1259,1300,1340,1374,1408,1442,1475,1509,1552,1595,1638,1681,1723,1749,1775,1801,1827,1853,1869,1886,1902,1919,1935,1944,1952,1961,1970,1978,1984,1989,1994,1999,2004,2007,2009,2012,2014,2017,2020,2023,2026,2029,2032,2033,2035,2037,2037,2037,2033,2036,2036]
}

RD100 = {
    'x': [0,0.5,1,1.5,2,2.5,3,3.5,4,4.5,5,5.5,6,6.5,7,7.5,8,8.5,9,9.5,10,10.5,11,11.5,12,12.5,12.9,13.5,14,14.4,15.1,15.6,16,16.5,17.1,17.5,18,18.5,19,19.5,20],
    'y': [0,0,0,0,0,0,5.3,63.1,112.8,166.9,245.4,327.4,450.1,594.8,761.7,939.6,1117.6,1330.7,1580,1759.2,1909.6,1980.4,2003.5,2009.8,2033.4,2033.3,2036.5,2033.4,2035.7,2036.6,2036.8,2036.5,2036.6,2038.2,2038.2,2038.2,2038.2,2038.2,2038.2,2038.2,2038.2]
}

RD113 = {
    'x': [0,0.5,1,1.5,2,2.5,3,3.5,4,4.5,5,5.5,6,6.5,7,7.5,8,8.5,9,9.5,10,10.5,11,11.5,12,12.5,13,13.5,14,14.5,15,15.5,16,16.5,17,17.5,18,18.5,19,19.5,20],
    'y': [0,0,0,0,0,0,21.6,74.5,139.6,216.8,311.8,427.1,563.8,735.6,933.1,1152.2,1358.95,1590.14,1866.4,1935.73,2001.8,2000.5,2000.6,2000.6,2001.8,2008.2,2010.8,2010.5,2011,2011,2011,2011,2011,2011,2011,2011,2011,2011,2011,2011,2011]
}

def load_data(file):
    """Load Excel data and return DataFrame"""
    try:
        df = pd.read_excel(file)
        return df
    except Exception as e:
        st.error(f"Error loading file: {str(e)}")
        return None

def generate_turbine_plot(turbine_data, turbine_name, model_name):
    """Generate power curve plot for a single turbine"""
    fig, ax = plt.subplots(figsize=(12, 8))
    
    # Filter valid data (remove NaN values)
    valid_data = turbine_data.dropna(subset=['Wind speed - AVE [m/s]', 'Active power - AVE [kW]'])
    
    # Extract Site, Customer, Week information from the data if available
    site = turbine_data['Site'].iloc[0] if 'Site' in turbine_data.columns else 'N/A'
    Customer = turbine_data['Customer'].iloc[0] if 'Customer' in turbine_data.columns else 'N/A'
    week = turbine_data['Week'].iloc[0] if 'Week' in turbine_data.columns else 'N/A'
    try :
        week = int(week)
    except :
        pass
    
    if len(valid_data) > 0:
        # Color mapping based on 'Power curve validity - MIN' values
        color_map = {0: 'yellow', 1: 'orange', 2: 'red', 3: 'blue'}  # 0 and 3 both map to blue
        
        if 'Power curve validity - MIN' in valid_data.columns:
            # Group data by validity values and plot with different colors
            for validity_value in valid_data['Power curve validity - MIN'].unique():
                if not pd.isna(validity_value):
                    subset = valid_data[valid_data['Power curve validity - MIN'] == validity_value]
                    color = color_map.get(validity_value, 'gray')  # Default to gray if value not in map
                    ax.scatter(subset['Wind speed - AVE [m/s]'], 
                              subset['Active power - AVE [kW]'], 
                              alpha=0.6, s=20, 
                              label=f'{turbine_name} (Validity: {validity_value})', 
                              color=color)
        else:
            # Default blue color if validity column doesn't exist
            ax.scatter(valid_data['Wind speed - AVE [m/s]'], 
                      valid_data['Active power - AVE [kW]'], 
                      alpha=0.6, s=20, label=f'{turbine_name} Actual Data', color='blue')
    
    # Plot model curve based on turbine model
    if model_name == 'RD93':
        ax.plot(RD93['x'], RD93['y'], 'r-', linewidth=2, label='RD93 Standard Curve')
    elif model_name == 'RD100':
        ax.plot(RD100['x'], RD100['y'], 'g-', linewidth=2, label='RD100 Standard Curve')
    elif model_name == 'RD113':
        ax.plot(RD113['x'], RD113['y'], 'm-', linewidth=2, label='RD113 Standard Curve')
    
    # Customize plot
    ax.set_xlabel('Wind Speed [m/s]', fontsize=12)
    ax.set_ylabel('Active Power [kW]', fontsize=12)
    
    # Updated title with Site, Customer, Week information
    title = f'Power Curve - {turbine_name} ({model_name})\nSite: {site} | Customer: {Customer} | Week: {week}'
    ax.set_title(title, fontsize=14, fontweight='bold')
    
    # Custom grid lines
    # X-axis: 0, 1, 2, 3, 4, ..., 20
    x_ticks = list(range(0, 18))
    ax.set_xticks(x_ticks)
    
    # Y-axis: 0, 100, 200, 300, 400, ..., 2100
    y_ticks = list(range(0, 2101, 100))
    ax.set_yticks(y_ticks)
    
    # Enable grid
    ax.grid(True, alpha=0.3, linestyle='-', linewidth=0.5)
    ax.legend()
    
    # Set axis limits
    ax.set_xlim(0, 20)
    ax.set_ylim(0, 2100)
    
    plt.tight_layout()
    return fig

def create_zip_file(plots_data, turbine_metadata):
    """Create a ZIP file containing all PNG files with Site and Week in filename"""
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for turbine_name, fig in plots_data.items():
            # Get site and week info for filename
            site = turbine_metadata[turbine_name]['site']
            week = turbine_metadata[turbine_name]['week']
            Customer = turbine_metadata[turbine_name]['Customer']
            try :
                week = int(week)
            except :
                pass
            # Create filename with Site and Week
            filename = f"{site}_{Customer}_{turbine_name}_Week{week}.png"
            
            # Save plot to bytes
            img_buffer = io.BytesIO()
            fig.savefig(img_buffer, format='png', dpi=300, bbox_inches='tight')
            img_buffer.seek(0)
            
            # Add to zip
            zip_file.writestr(filename, img_buffer.getvalue())
            plt.close(fig)  # Close figure to free memory
    
    zip_buffer.seek(0)
    return zip_buffer

def main():
    st.set_page_config(page_title="Turbine Power Curve Analysis", layout="wide")
    
    st.title("Inox Turbine Power Curve Analysis")
    st.markdown("Upload your turbine data Excel file to generate individual power curve PNG files for each turbine.")
    
    # Sidebar for file upload
    st.sidebar.header("📁 File Upload")
    uploaded_file = st.sidebar.file_uploader("Choose Excel file", type=['xlsx', 'xls'])
    
    # Model information
    st.sidebar.header("📊 Supported Models")
    st.sidebar.info("RD93, RD100, RD113")
    
    # # Color coding information
    # st.sidebar.header("🎨 Color Coding")
    # st.sidebar.info("Blue: Power curve validity 0 & 3\nYellow: Power curve validity 1\nOrange: Power curve validity 2")
    
    # Load sample data if no file uploaded
    df = None
    if uploaded_file is None:
        st.sidebar.info("No file uploaded. Using sample data for demonstration.")
        try:
            df = pd.read_excel("data.xlsx")
            st.success("✅ Sample data loaded successfully!")
        except Exception as e:
            st.warning("⚠️ No sample data available. Please upload an Excel file.")
            st.info("Expected Excel format: Columns should include 'Turbine', 'Model', 'Site', 'Customer', 'Week', 'Wind speed - AVE [m/s]', 'Active power - AVE [kW]', 'Power curve validity - MIN'")
            return
    else:
        df = load_data(uploaded_file)
        if df is None:
            return
        st.success("✅ File uploaded successfully!")
    
    # Display data info
    st.header("📊 Data Overview")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("Total Records", len(df))
    
    with col2:
        unique_turbines = df['Turbine'].nunique()
        st.metric("Unique Turbines", unique_turbines)
    
    with col3:
        unique_models = df['Model'].nunique()
        st.metric("Models", unique_models)
    
    # Show data sample
    st.subheader("Data Sample")
    st.dataframe(df.head(10))
    
    # Turbine analysis
    st.header("🔧 Turbine Analysis")
    
    # Get unique turbines
    turbines = df['Turbine'].unique()
    
    if len(turbines) == 0:
        st.error("No turbines found in the data.")
        return
    
    # Display turbine list
    st.subheader("Available Turbines")
    turbine_df = df.groupby(['Turbine', 'Model']).size().reset_index(name='Records')
    st.dataframe(turbine_df)
    
    # Generate plots button
    if st.button("🚀 Generate Power Curve Plots", type="primary"):
        st.header("📈 Generated Plots")
        
        progress_bar = st.progress(0)
        plots_data = {}
        turbine_metadata = {}
        
        for i, turbine in enumerate(turbines):
            # Get turbine data
            turbine_data = df[df['Turbine'] == turbine]
            model_name = turbine_data['Model'].iloc[0]
            
            # Store metadata for filename generation
            site = turbine_data['Site'].iloc[0] if 'Site' in turbine_data.columns else 'N/A'
            week = turbine_data['Week'].iloc[0] if 'Week' in turbine_data.columns else 'N/A'
            Customer = turbine_data['Customer'].iloc[0] if 'Customer' in turbine_data.columns else 'N/A'
            turbine_metadata[turbine] = {'site': site, 'week': week,'Customer':Customer }
            
            # Generate plot
            fig = generate_turbine_plot(turbine_data, turbine, model_name)
            plots_data[turbine] = fig
            
            # Update progress
            progress_bar.progress((i + 1) / len(turbines))
        
        st.success(f"✅ Generated {len(plots_data)} power curve plots!")
        
        # Display sample plots
        st.subheader("Sample Plots Preview")
        sample_turbines = list(plots_data.keys())[:3]  # Show first 3 plots
        
        for turbine in sample_turbines:
            st.subheader(f"Turbine: {turbine}")
            st.pyplot(plots_data[turbine])
        
        if len(plots_data) > 3:
            st.info(f"Showing 3 of {len(plots_data)} plots. Download the ZIP file to get all plots.")
        
        # Create download button
        st.header("📥 Download Results")
        
        zip_buffer = create_zip_file(plots_data, turbine_metadata)
        
        st.download_button(
            label="📦 Download All PNG Files (ZIP)",
            data=zip_buffer.getvalue(),
            file_name=f"turbine_power_curves_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
            mime="application/zip"
        )
        
        st.success("🎉 All plots generated successfully! Click the download button to get your PNG files with Site and Week information in filenames.")

if __name__ == "__main__":
    main()