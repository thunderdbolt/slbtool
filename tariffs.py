import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl.styles import Font

# -----------------------------------------------------------------------------
# Streamlit Tab and Page Configuration
# -----------------------------------------------------------------------------

#Tab name config
st.set_page_config(page_icon="https://upload.wikimedia.org/wikipedia/commons/thumb/d/d6/SLB_Logo_2022.svg/2560px-SLB_Logo_2022.svg.png", page_title="Tariffs Processing", layout="wide")
# URLs for the images
slb_logo_url = "https://upload.wikimedia.org/wikipedia/commons/thumb/d/d6/SLB_Logo_2022.svg/2560px-SLB_Logo_2022.svg.png"
slb_website_url = "https://www.slb.com/"  # The URL you want to open when the logo is clicked
pipesim_logo_url = "https://americanpatriotsdecals.com/cdn/shop/products/American-Bald-Eagle-American-Flag.jpg?v=1638310152"
# pipesim_website_url = "https://www.software.slb.com/products/pipesim"
mid_image_url = "https://github.com/thunderdbolt/slbtool/blob/main/new.jpg?raw=true"
seal_award_image_url = "https://cei.org/wp-content/uploads/2017/08/2000px-Seal_of_the_United_States_Department_of_Commerce.svg_.png"
seal_award_website_url = "https://dataweb.usitc.gov/"
ghg_image_url = "https://github.com/thunderdbolt/slbtool/blob/main/tariffs.png?raw=true"

# Custom CSS to remove padding and margin from columns
st.markdown(
    """
    <style>
    .css-1e5imcs {padding: 0!important;}
    .st-br {margin: 0!important;}
    </style>
    """,
    unsafe_allow_html=True,
)

# Display images with responsive design
col1, col2, col3, col4 = st.columns([1, 1, 6, 2], gap="small")  # Adjust column ratios as needed

# Determine the max height based on your layout requirements or through trial and error
max_image_height = "100px"  # Adjust this value as needed to fit your header layout

with col1:
    st.markdown(
        f'<a href="{slb_website_url}">'
        f'<img src="{slb_logo_url}" alt="SLB Logo" style="max-height:{max_image_height}; height:auto; width:auto;">'
        f'</a>',
        unsafe_allow_html=True,
    )

with col2:
    st.markdown(
        # f'<a href="{pipesim_website_url}">'
        f'<img src="{pipesim_logo_url}" alt="Pipesim Logo" style="max-height:{max_image_height}; height:auto; width:auto;">'
        f'</a>',
        unsafe_allow_html=True,
    )

with col3:
    st.markdown(
        # f'<a href="{slb_website_url}">'
        f'<img src="{mid_image_url}" alt="Middle Image" style="max-height:{max_image_height}; width:93%;">'
        f'</a>',
        unsafe_allow_html=True,
    )
    
with col4:
    st.markdown(
        f'<a href="{seal_award_website_url}">'
        f'<img src="{seal_award_image_url}" alt="Seal Award Logo" style="max-height:{max_image_height}; height:auto; width:auto;">'
        f'</a>',
        unsafe_allow_html=True,
    )




###############################################################################################################################################
# Functions Being Used
###############################################################################################################################################
        
def clean_hts_code(hts_code):
    # Convert to string and remove all dots and whitespace
    hts_code_str = str(hts_code)
    return hts_code_str.replace('.', '').strip()
#==============================================================================================
def reformat_hts_code(hts_code):
    # Reformat to xxxx.xx.xx.xx or xxxx.xx.xxxx depending on length
    # First ensure it's a string without any whitespace or dots
    hts_code_str = clean_hts_code(hts_code)
    if len(hts_code_str) > 10:
        return f"{hts_code_str[:4]}.{hts_code_str[4:6]}.{hts_code_str[6:8]}.{hts_code_str[8:]}"
    else:
        return f"{hts_code_str[:4]}.{hts_code_str[4:6]}.{hts_code_str[6:]}"
#==============================================================================================
@st.cache_data
def process_sheet(df, label):
    df = df[df['HSCODE'].notna()]
    df['ADD/CVD'] = label
    df['HSCODE'] = df['HSCODE'].apply(clean_hts_code)
    df = df.rename(columns={'HSCODE': 'HTS_Code'})[['HTS_Code', 'ADD/CVD']]
    return df
#==============================================================================================
@st.cache_data
def combine_data(file):
    sheet_names = pd.ExcelFile(file).sheet_names
    combined_data = []

    for sheet in sheet_names:
        df = pd.read_excel(file, sheet_name=sheet)
        if 'ADD' in sheet:
            combined_data.append(process_sheet(df, 'ADD'))
        elif 'CVD' in sheet:
            combined_data.append(process_sheet(df, 'CVD'))

#==============================================================================================
@st.cache_data
def find_hts_column(df):
    possible_columns = ['HTS', 'HTS Number', 'HSCODE']
    for col in possible_columns:
        if col in df.columns:
            return col
    raise ValueError("HTS column not found in the dataframe.")
#==============================================================================================
@st.cache_data
def process_rate_of_duty(file):
    df = pd.read_excel(file)
    hts_col_name = find_hts_column(df)
    df[hts_col_name] = df[hts_col_name].apply(clean_hts_code)
    df = df.rename(columns={hts_col_name: 'HTS_Code', 'General Rate of Duty': 'General_Rate_of_Duty'})
    return df
#==============================================================================================
@st.cache_data
def process_specific_duty(file, column_name):
    # This function processes the specific duty rates for China, Aluminum, or Steel.
    df = pd.read_excel(file)
    df['HTS'] = df['HTS'].apply(clean_hts_code)
    df = df.rename(columns={'HTS': 'HTS_Code', column_name: 'Specific_Rate_of_Duty'})
    df['Specific_Rate_of_Duty'] = df['Specific_Rate_of_Duty'].fillna(0)  # Assuming empty cells mean 0%
    return df
#==============================================================================================
@st.cache_data
def merge_all_data(uploaded_files):
    addcvd_dataframes = []
    specific_duty_dataframes = []
    rate_of_duty_dataframes = []

    for uploaded_file in uploaded_files:
        # Determine the type of file
        sheet_names = pd.ExcelFile(uploaded_file).sheet_names
        is_specific_duty = 'China' in sheet_names or 'Aluminum' in sheet_names or 'Steel' in sheet_names
        is_addcvd = any(['ADD' in sheet for sheet in sheet_names]) or any(['CVD' in sheet for sheet in sheet_names])
        
        if is_addcvd:
            addcvd_dataframes.append(combine_data(uploaded_file))
        elif is_specific_duty:
            specific_duty_dataframes.append(process_specific_duty(uploaded_file, sheet_names[0]))
        else:
            rate_of_duty_dataframes.append(process_rate_of_duty(uploaded_file))

    # Initialize merged_data as an empty DataFrame or the first non-empty DataFrame
    merged_data = pd.DataFrame()

    # Merge all dataframes of each type
    if addcvd_dataframes:
        combined_addcvd = pd.concat(addcvd_dataframes, ignore_index=True)
        merged_data = combined_addcvd

    if specific_duty_dataframes:
        combined_specific_duty = pd.concat(specific_duty_dataframes, ignore_index=True)
        if not merged_data.empty:
            merged_data = pd.merge(merged_data, combined_specific_duty, on='HTS_Code', how='left')
        else:
            merged_data = combined_specific_duty

    if rate_of_duty_dataframes:
        combined_rate_of_duty = pd.concat(rate_of_duty_dataframes, ignore_index=True)
        if not merged_data.empty:
            merged_data = pd.merge(merged_data, combined_rate_of_duty, on='HTS_Code', how='left')
        else:
            merged_data = combined_rate_of_duty

    return merged_data
#==============================================================================================
def combine_tariff_information(row):
    # Combine values from General_Rate_of_Duty, Steel, Aluminum, and China columns
    # Convert each item to a string and use a list comprehension to filter out None or empty string values before joining
    values_to_combine = [str(row['General_Rate_of_Duty']), str(row.get('Steel')), str(row.get('Aluminum')), str(row.get('China'))]
    combined_info = ' + '.join([v for v in values_to_combine if v and v != 'nan'])
    return combined_info


###############################################################################################################################################
# Streamlit Application Logic
###############################################################################################################################################

def main():
    st.title("HTS Code Data Processing and Tariff Calculation Tool")

    # Streamlit app layout with tabs
    tab1, tab2, tab3 = st.tabs(["Welcome", "Data Entry", "Data Description"])

    with tab1:
        # st.write("")
        col1, col2 = st.columns([1,1])

        with col1:
            st.write("""
            This tool is designed to provide a comprehensive solution for calculating accurate tariffs on international goods. With an intuitive interface, users can easily input key details such as HTS Codes, country of origin, transportation method, and the value and weight of items, streamlining the process of tariff calculation.

            ## How to Use
            1. **Enter Details**: Navigate to the 'Data Entry' tab to input essential information, including HTS Codes, country of origin, value, weight, and transportation method.
            2. **Calculate Tariffs**: Automatically computes the tariffs based on your inputs, taking into account various factors such as general tariff percentages and specific duties.
            3. **Review Results**: Offers a detailed view of the calculated tariffs and fees, providing a comprehensive understanding of potential costs.
            4. **Adjust and Recalculate**: Enables modification of input values to explore different scenarios and immediately see the updated tariffs.
            5. **Export Data**: Allows for exporting the detailed tariff calculations for further analysis or record-keeping.
            """)

        with col2:
            st.markdown(
                # f'<a href="{slb_website_url}">'
                f'<img src="{ghg_image_url}" alt="SLB Logo" style="max-height:auto; width:auto; display: block; margin-left: auto; margin-right: auto;"></a>',
                unsafe_allow_html=True,
            )

            # # Embed a YouTube video with a specific size
            # video_url = 'https://www.youtube.com/embed/40PaXCJfi6M'  # Note: Use the 'embed' URL
            # st.markdown(f'<iframe width="800" height="410" src="{video_url}" frameborder="0" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe>', unsafe_allow_html=True)


        slb_logo_explained = "https://user-images.githubusercontent.com/135745865/255714845-eb74d101-c6d8-4dc7-9ce5-fa791593ddbf.jpeg"
        slb_logo_explained_2 = "https://user-images.githubusercontent.com/135745865/255726767-c571acdc-4a22-4859-b196-892bf5f622d3.png"
        slb_logo_explained_url = "https://www.slb.com/about/who-we-are/for-a-balanced-planet"

        st.divider()
        st.write("")
        st.write("")
        st.write("")
        st.write("") 
        col1, col2, col3, col4 = st.columns([3, 1, 2, 1])
        with col1:
            # st.markdown(
            #     f'<a href="{slb_logo_explained_url}"><img src="{slb_logo_explained}" alt="SLB Logo 6" style="width:800px; height:600px;"></a><p style="text-align:center;">HELLOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO</p>',
            #     unsafe_allow_html=True,
            # )
            st.image(slb_logo_explained, caption="This is the reduction in carbon emissions the world needs in order to meet its net-zero objectives by 2050, but it is also what we need to do to maintain this for the long term. We need to balance the impact of the past with sustained net negative carbon solutions. We placed this global energy challenge at the heart of what we do and who we are. It’s the path we have all taken, and we will support its realization.", width=720)


        with col3:
            # st.markdown(
            #     f'<a href="{slb_logo_explained_url}"><img src="{slb_logo_explained_2}" alt="SLB Logo 7" style="width:800px; height:600px;"></a>',
            #     unsafe_allow_html=True,
            # )
            st.write("")
            st.write("")
            st.image(slb_logo_explained_2, caption="The SLB logo represents our bold commitment, not only to our own net-zero journey, but also in support of our customers’, while embodying our pledge to go beyond. And all to bring balance back to our planet.", width=660)


        st.markdown("""
            <style>
            .footer {
                # position: fixed;
                left: 0;
                bottom: 0;
                width: 100%;
                text-align: center;
                color: #888;
                font-size: 12px;
                margin-top: 3em;  # Add space before the text
            }
            </style>
            <div class='footer'>
                <p>Copyright © 2024 SLB. All rights reserved.</p>
            </div>
        """, unsafe_allow_html=True)
        st.divider()

    with tab2:
        # Unified file uploader for both ADD/CVD and Rate of Duty files
        st.subheader("Upload your Excel files")
        uploaded_files = st.file_uploader("Upload", type=["xls", "xlsx"], key="file_uploader", accept_multiple_files=True)

    if uploaded_files:
        # Process uploaded files
        for uploaded_file in uploaded_files:
            # Check the type of file by reading the sheets or some other method
            sheet_names = pd.ExcelFile(uploaded_file).sheet_names
            is_specific_duty = 'China' in sheet_names or 'Aluminum' in sheet_names or 'Steel' in sheet_names
            is_addcvd = any(['ADD' in sheet for sheet in sheet_names]) or any(['CVD' in sheet for sheet in sheet_names])
            
            if is_addcvd:
                # Process as ADD/CVD file
                st.session_state['combined_addcvd_data'] = combine_data(uploaded_file)
                with tab3:
                    # if st.checkbox(f'Show ADD/CVD Data for {uploaded_file.name}'):
                        with st.expander(f"Data from {uploaded_file.name}:"):
                            st.dataframe(st.session_state['combined_addcvd_data'], use_container_width=True)

            elif is_specific_duty:
                # Process as specific duty rate file
                st.session_state['specific_duty_data'] = process_specific_duty(uploaded_file, sheet_names[0])
                with tab3:
                    # if st.checkbox(f'Show {sheet_names[0]} Data for {uploaded_file.name}'):
                        with st.expander(f"Data from {uploaded_file.name}:"):
                            st.dataframe(st.session_state['specific_duty_data'], use_container_width=True)

            else:
                # Process as Rate of Duty file
                st.session_state['rate_of_duty_data'] = process_rate_of_duty(uploaded_file)
                with tab3:
                #    if st.checkbox(f'Show General Rate of Duty Data for {uploaded_file.name}'):
                        with st.expander(f"Data from {uploaded_file.name}:"):
                            st.dataframe(st.session_state['rate_of_duty_data'], use_container_width=True)

    with tab2:
        if uploaded_files:
            processed_data = merge_all_data(uploaded_files)
            try:
                processed_data['Tariff'] = processed_data.apply(combine_tariff_information, axis=1)
            except Exception as e:
                st.error(f"Error processing data: {e}")
            st.write("Processed Data (with Tariffs):")
            with st.expander('HTS Data', expanded=True):
                # st.dataframe(processed_data, use_container_width=True)

                # Assuming this function looks up tariff information based on the HTS code
                def lookup_tariff_data(hts_code, processed_data):
                    # Assuming there's a column 'China Tariff' in your processed_data
                    tariff_info = processed_data[processed_data['HTS_Code'] == hts_code]
                    if not tariff_info.empty:
                        # Extract the specific tariff data from the dataframe
                        return {
                            'General_Rate_of_Duty': tariff_info['General_Rate_of_Duty'].iloc[0],
                            'China Duties': tariff_info['China Duties'].iloc[0],
                            'Aluminum': tariff_info['Aluminum'].iloc[0],
                            'Steel': tariff_info['Steel'].iloc[0],
                            'ADD/CVD Flag': tariff_info['ADD/CVD Flag'].iloc[0],
                        }
                    return {'China Duties': 0, 'Aluminum': 0, 'Steel': 0, 'ADD/CVD Flag': ''}

                # Create an editable table
                def display_editable_table():
                    # Define the structure of the empty DataFrame according to the Excel sheet
                    data_structure = {
                        "SLB Part Number": [""],
                        "US HTS": [""],
                        "COO": [""],
                        "Value": [100],
                        "Weight": [0],
                        "MOT": [""],
                        "General Tariff Percentage": [0],  # To be calculated based on US HTS
                        "COO China Tariff": [0],  # To be calculated based on COO
                        "Aluminum Tariff": [0],  # To be calculated based on US HTS
                        "Steel Tariff": [0],  # To be calculated based on US HTS
                        "Potential ADD/CVD Flag": [""],  # To be determined based on US HTS and COO
                        "CBP Merchandise Processing Fee": [31.67],  # Fixed cost for all items
                        "CBP Harbor Maintenance Fee": [0],  # To be calculated based on Value and MOT
                        "Tariffs & Fees to be Paid (%)": [0],  # To be calculated based on above tariffs + 0.35%
                        "Tariffs to be Paid (USD)": [0],  # To be calculated based on Value and the total percentage
                        "Tariffs & Fees to be Paid (USD)": [0],  # To include the $31.67 fee in the total
                    }
                    df = pd.DataFrame(data_structure)

                    # Define column configurations with a dropdown for COO
                    column_config = {
                        "COO": st.column_config.SelectboxColumn(
                            options=[
                                "China", "USA", "Italy", "Germany", "France",
                                "Afghanistan", "Albania", "Algeria", "Andorra", "Angola",
                                "Argentina", "Armenia", "Australia", "Austria", "Azerbaijan",
                                "Bahamas", "Bahrain", "Bangladesh", "Barbados", "Belarus",
                                "Belgium", "Belize", "Benin", "Bhutan", "Bolivia",
                                "Bosnia and Herzegovina", "Botswana", "Brazil", "Brunei", "Bulgaria",
                                "Burkina Faso", "Burundi", "Cambodia", "Cameroon", "Canada",
                                "Cape Verde", "Central African Republic", "Chad", "Chile", "Colombia",
                                "Comoros", "Congo", "Costa Rica", "Croatia", "Cuba",
                                "Cyprus", "Czech Republic", "Denmark", "Djibouti", "Dominica",
                                "Dominican Republic", "East Timor", "Ecuador", "Egypt", "El Salvador",
                                "Equatorial Guinea", "Eritrea", "Estonia", "Eswatini", "Ethiopia",
                                "Fiji", "Finland", "Gabon", "Gambia", "Georgia",
                                "Ghana", "Greece", "Grenada", "Guatemala", "Guinea",
                                "Guinea-Bissau", "Guyana", "Haiti", "Honduras", "Hungary",
                                "Iceland", "India", "Indonesia", "Iran", "Iraq",
                                "Ireland", "Israel", "Jamaica", "Japan", "Jordan",
                                "Kazakhstan", "Kenya", "Kiribati", "Kosovo", "Kuwait",
                                "Kyrgyzstan", "Laos", "Latvia", "Lebanon", "Lesotho",
                                "Liberia", "Libya", "Liechtenstein", "Lithuania", "Luxembourg",
                                "Madagascar", "Malawi", "Malaysia", "Maldives", "Mali",
                                "Malta", "Marshall Islands", "Mauritania", "Mauritius", "Mexico",
                                "Micronesia", "Moldova", "Monaco", "Mongolia", "Montenegro",
                                "Morocco", "Mozambique", "Myanmar", "Namibia", "Nauru",
                                "Nepal", "Netherlands", "New Zealand", "Nicaragua", "Niger",
                                "Nigeria", "North Korea", "North Macedonia", "Norway", "Oman",
                                "Pakistan", "Palau", "Palestine", "Panama", "Papua New Guinea",
                                "Paraguay", "Peru", "Philippines", "Poland", "Portugal",
                                "Qatar", "Romania", "Russia", "Rwanda", "Saint Kitts and Nevis",
                                "Saint Lucia", "Saint Vincent and the Grenadines", "Samoa", "San Marino", "Sao Tome and Principe",
                                "Saudi Arabia", "Senegal", "Serbia", "Seychelles", "Sierra Leone",
                                "Singapore", "Slovakia", "Slovenia", "Solomon Islands", "Somalia",
                                "South Africa", "South Korea", "South Sudan", "Spain", "Sri Lanka",
                                "Sudan", "Suriname", "Sweden", "Switzerland", "Syria",
                                "Taiwan", "Tajikistan", "Tanzania", "Thailand", "Togo",
                                "Tonga", "Trinidad and Tobago", "Tunisia", "Turkey", "Turkmenistan",
                                "Tuvalu", "Uganda", "Ukraine", "United Arab Emirates", "United Kingdom",
                                "Uruguay", "Uzbekistan", "Vanuatu", "Vatican City", "Venezuela",
                                "Vietnam", "Yemen", "Zambia", "Zimbabwe"
                            ],


                            help="Select the Country of Origin"
                        ),
                        "MOT": st.column_config.SelectboxColumn(
                            options=[
                                "AIR", "TRUCK", "OCEAN", "COURIER",
                            ],


                            help="Select the Method of Transportation"
                        ),
                    }

                    # Assuming df is your initial dataframe loaded with data
                    input_columns = ["SLB Part Number", "US HTS", "COO", "Value", "Weight", "MOT"]
                    outcome_columns = ["General Tariff Percentage", "COO China Tariff", "Aluminum Tariff", "Steel Tariff", 
                                    "Potential ADD/CVD Flag", "CBP Merchandise Processing Fee", "CBP Harbor Maintenance Fee", 
                                    "Tariffs & Fees to be Paid (%)", "Tariffs to be Paid (USD)", "Tariffs & Fees to be Paid (USD)"]


                    # Display the editable table with specified configurations
                    col1, col2 = st.columns([3,7], gap="small")
                    with col1:
                        new_df = st.data_editor(
                        df[input_columns],  # Show only input columns for editing
                        num_rows="dynamic",
                        column_config=column_config,
                        key="data_editor"
                        )

                    # Apply the cleaning function to the HTS codes column after the user input
                    if 'US HTS' in new_df.columns:
                        new_df['US HTS'] = new_df['US HTS'].apply(clean_hts_code)

                    # After editing, you can combine new_df with the outcome columns to get the full table
                    full_df = pd.concat([new_df, df[outcome_columns]], axis=1)

                    # Ensure CBP Merchandise Processing Fee is set for all rows
                    full_df['CBP Merchandise Processing Fee'] = full_df['CBP Merchandise Processing Fee'].fillna(31.67)

                    st.session_state['user_data'] = full_df
                    st.session_state['new_df'] = new_df

                    def parse_general_rate(rate, value, weight):
                        # Convert rate to string if it's not already a string
                        rate_str = str(rate)
                        
                        try:
                            if '%' in rate_str:
                                # If the rate is a simple percentage
                                return float(rate_str.replace('%', '')) / 100
                            elif '/kg' in rate_str or '/liter' in rate_str:
                                # If the rate is a currency per weight or per liter
                                rate_value = float(rate_str.split('/')[0].replace('¢', '').replace('$', ''))
                                if '¢' in rate_str:
                                    rate_value /= 100  # Convert cents to dollars
                                # Return total cost based on weight or volume
                                return (rate_value * weight) / value
                            elif '+' in rate_str:
                                # If the rate has a fixed part and a percentage part, like "$1.035/kg + 17%"
                                fixed_part, percentage_part = rate_str.split(' + ')
                                fixed_rate = parse_general_rate(fixed_part, value, weight)
                                percentage_rate = parse_general_rate(percentage_part, value, weight)
                                return fixed_rate + percentage_rate
                            elif '$' in rate_str:
                                # If the rate is a fixed currency amount (assumed per kg for this calculation)
                                return (float(rate_str.replace('$', '')) * weight) / value
                            else:
                                # If the rate is already a float, assume it's a direct percentage
                                return float(rate) / 100
                        except ValueError:
                            return 0


                    # Inside your 'Update Tariffs' checkbox block
                    # if st.checkbox('Update Tariffs', value=True):  
                    Total_Tariffs = 0  # Initialize total tariffs sum
                    for index, row in new_df.iterrows():
                        hts_code = row['US HTS']
                        weight = row.get('Weight', 0)  # Ensure there is a default weight value
                        value = row.get('Value', 1)  # Ensure there is a default value

                        if hts_code:
                            tariff_info = lookup_tariff_data(hts_code, processed_data)
                            general_rate = tariff_info.get('General_Rate_of_Duty', '0%')
                            general_tariff_percentage = parse_general_rate(general_rate, value, weight)
                            
                            # Initialize tariffs to 0
                            china_duties = 0
                            aluminum_tariff = 0
                            steel_tariff = 0

                            # general_tariff_percentage = parse_tariff(tariff_info.get('General_Rate_of_Duty', '0%'), row.get('Weight', 0))

                            # Apply tariffs based on the conditions
                            if row['COO'] == 'China':
                                china_duties = float(tariff_info.get('China Duties', 0) or 0)
                            aluminum_tariff = float(tariff_info.get('Aluminum Tariff', 0) or 0)
                            steel_tariff = float(tariff_info.get('Steel Tariff', 0) or 0)

                            # Update the dataframe with the retrieved tariff information
                            new_df.at[index, 'General Tariff Percentage'] = general_tariff_percentage*10000
                            new_df.at[index, 'COO China Tariff'] = china_duties*100
                            new_df.at[index, 'Aluminum Tariff'] = aluminum_tariff
                            new_df.at[index, 'Steel Tariff'] = steel_tariff
                            new_df.at[index, 'Potential ADD/CVD Flag'] = tariff_info.get('ADD/CVD Flag', '')

                            # Calculate 'Tariffs & Fees to be Paid (%)'
                            total_percentage = (general_tariff_percentage + china_duties + aluminum_tariff + steel_tariff + 31.67/value)*100
                            new_df.at[index, 'Tariffs & Fees to be Paid (%)'] = total_percentage

                            # Calculate 'Tariffs to be Paid (USD)'
                            total_tariffs = (general_tariff_percentage + china_duties + aluminum_tariff + steel_tariff)*value
                            new_df.at[index, 'Tariffs to be Paid (USD)'] = total_tariffs

                            # Calculate 'Tariffs & Fees to be Paid (USD)'
                            total_percentage_usd = (general_tariff_percentage + china_duties + aluminum_tariff + steel_tariff)*value + 31.67
                            new_df.at[index, 'Tariffs & Fees to be Paid (USD)'] = total_percentage_usd
                            
                            Total_Tariffs += new_df.at[index, 'Tariffs & Fees to be Paid (USD)']

                    try:  
                        if 'CBP Merchandise Processing Fee' not in new_df.columns:
                            new_df['CBP Merchandise Processing Fee'] = 31.67

                        if 'CBP Harbor Maintenance Fee' not in new_df.columns:
                            new_df['CBP Harbor Maintenance Fee'] = 0

                        # Now define the outcome columns for display
                        outcome_columns = [
                            "General Tariff Percentage", "COO China Tariff", "Aluminum Tariff", "Steel Tariff", 
                            "Potential ADD/CVD Flag", "CBP Merchandise Processing Fee", "CBP Harbor Maintenance Fee", 
                            "Tariffs & Fees to be Paid (%)", "Tariffs to be Paid (USD)", "Tariffs & Fees to be Paid (USD)"
                        ]

                        # Filter 'new_df' to get only the outcome columns for display
                        outcome_df = new_df[outcome_columns]

                        # Update session state if needed
                        st.session_state['Total_Tariffs'] = Total_Tariffs
                        st.session_state['editable_data'] = new_df

                        # Display the updated tariff information
                        # st.write("Updated Tariff Information:")
                        with col2:
                            st.dataframe(outcome_df, use_container_width=True, hide_index=True)

                        # Display the total tariffs and fees
                        st.write(f"**Total Tariffs & Fees to be Paid (USD): {Total_Tariffs}**")

                        # # Display the total tariffs and fees with HTML for styling
                        # total_tariffs_html = f"""
                        #     <div style="
                        #         color: red; 
                        #         font-size: 24px; 
                        #         text-align: center;
                        #     ">
                        #         <strong>Total Tariffs & Fees to be Paid (USD): {Total_Tariffs}</strong>
                        #     </div>
                        # """
                        # st.markdown(total_tariffs_html, unsafe_allow_html=True)

                        # -----------------------------------------------------------------------------
                        # Function to convert and download dataframe to Excel
                        # -----------------------------------------------------------------------------
        
                        # Define the order of the columns as they should appear in the exported file
                        column_order = [
                            "SLB Part Number", "US HTS", "COO", "Value", "Weight", "MOT",
                            "General Tariff Percentage", "COO China Tariff", "Aluminum Tariff", "Steel Tariff",
                            "Potential ADD/CVD Flag", "CBP Merchandise Processing Fee", "CBP Harbor Maintenance Fee",
                            "Tariffs & Fees to be Paid (%)", "Tariffs to be Paid (USD)", "Tariffs & Fees to be Paid (USD)"
                        ]
        
                        # Ensure new_df columns are in the correct order
                        new_df = st.session_state['new_df'][column_order]
                        
                        # Calculate the total tariffs if not already done
                        Total_Tariffs = new_df['Tariffs & Fees to be Paid (USD)'].sum()
                        
                        # Add the Total_Tariffs to the new_df
                        total_row_data = [None]*(len(column_order)-1) + [Total_Tariffs]
                        total_row_df = pd.DataFrame([total_row_data], columns=column_order)
                        new_df = pd.concat([new_df, total_row_df], ignore_index=True)
        
                        
                        # Call the function to make the download button available in the Streamlit app
                        # Function to convert DataFrame to Excel
                        def to_excel(df):
                            # Create a BytesIO buffer to hold the Excel file in memory
                            output = BytesIO()
                            
                            # Create a Pandas Excel writer using the 'openpyxl' engine and the BytesIO buffer
                            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                # Write the DataFrame to the Excel writer
                                df.to_excel(writer, index=False, sheet_name='Sheet1')
                                
                                # Access the openpyxl workbook and worksheet objects to apply styles
                                workbook = writer.book
                                worksheet = writer.sheets['Sheet1']
                        
                                # Define the font style for the total row
                                bold_red_font = Font(bold=True, color="FF0000")
                                
                                # Get the max row (last row) in the worksheet
                                max_row = worksheet.max_row
                                
                                # Apply the font style to all cells in the last row
                                for row in worksheet.iter_rows(min_row=max_row, max_row=max_row):
                                    for cell in row:
                                        cell.font = bold_red_font
                            
                            # At this point, the ExcelWriter context is closed and the data is saved to the output buffer
                            return output.getvalue()
                        
                        def download_excel(df):
                            excel_data = to_excel(df)
                            st.download_button(label='📥 Download Excel',
                                               data=excel_data,
                                               file_name='tariff_data.xlsx',
                                               mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        
        
                        # Ensure new_df is available in the session state before attempting to download
                        if 'new_df' in st.session_state:
                            download_excel(new_df)
                        else:
                            st.error("No data available to download.")
                    
                    except:
                        pass    
                display_editable_table()     
                

if __name__ == "__main__":
    main()
