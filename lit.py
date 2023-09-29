import base64
import zipfile

import fitz
from operator import itemgetter
from itertools import groupby
import pandas as pd
import re
import os
import streamlit as st
import numpy as np

def ParseTab(page, bbox, columns=None):
    """Returns the parsed table of a page in a PDF / (open) XPS / EPUB document.
    Parameters:
    page: fitz.Page object
    bbox: containing rectangle, list of numbers [xmin, ymin, xmax, ymax]
    columns: optional list of column coordinates. If None, columns are generated
    Returns the parsed table as a list of lists of strings.
    The number of rows is determined automatically
    from parsing the specified rectangle.
    """
    tab_rect = fitz.Rect(bbox).irect
    xmin, ymin, xmax, ymax = tuple(tab_rect)

    if tab_rect.is_empty or tab_rect.is_infinite:
        return []

    if type(columns) is not list or columns == []:
        coltab = [tab_rect.x0, tab_rect.x1]
    else:
        coltab = sorted(columns)

    if xmin < min(coltab):
        coltab.insert(0, xmin)
    if xmax > coltab[-1]:
        coltab.append(xmax)

    words = page.get_text("words")

    if words == []:
        print("Warning: page contains no text")
        return []

    alltxt = []

    # get words contained in table rectangle and distribute them into columns
    for w in words:
        ir = fitz.Rect(w[:4]).irect  # word rectangle
        if ir in tab_rect:
            cnr = 0  # column index
            for i in range(1, len(coltab)):  # loop over column coordinates
                if ir.x0 < coltab[i]:  # word start left of column border
                    cnr = i - 1
                    break
            alltxt.append([ir.x0, ir.y0, ir.x1, cnr, w[4]])

    if alltxt == []:
        print("Warning: no text found in rectangle!")
        return []

    alltxt.sort(key=itemgetter(1))  # sort words vertically

    # create the table / matrix
    spantab = []  # the output matrix

    for y, zeile in groupby(alltxt, itemgetter(1)):
        schema = [""] * (len(coltab) - 1)
        for c, words in groupby(zeile, itemgetter(3)):
            entry = " ".join([w[4] for w in words])
            schema[c] = entry
        spantab.append(schema)

    return spantab


def extract_file_data(doc):

    w = doc[0].rect.width
    h = doc[0].rect.height

    def parse_futures():
        """
        xmin: 0 (Start from the left edge)
        ymin: 841 * 2/3 (Start 1/3 down from the top)
        xmax: 595 * 1/2 (Cover 1/2 the width of the page)
        ymax: 841 * 4/5 (End where the date starts. This could be adjusted as needed)

        :return:
        """
        return ParseTab(doc[0], [0, h *1/7, w, h * 46/100])

    return parse_futures()


def parse_to_table(data):

    # Initialize an empty list to store the parsed data
    parsed_data = []

    # Initialize the date
    date = None

    # Loop through the data
    for row in data:
        # Extract the date
        if 'as of' in row[0]:
            date = row[0].split('as of')[1].strip()
        # Parse the row if it starts with 'CBOT'
        elif row[0].startswith('CBOT'):
            # Split the string into parts
            parts = row[0].split()
            exchange = parts[0]
            commodity = parts[1]

            # Combine the remaining parts into a single string
            prices_str = ' '.join(parts[2:])

            # Use regex to find all matches of the pattern "price (month year)"
            matches = re.findall(r'(\d+\.\d+) \((\w+ \d+)\)', prices_str)

            # Loop through the matches and append the parts to the list
            for match in matches:
                price = match[0]
                month_year = match[1].split()
                month = month_year[0]
                year = month_year[1]
                parsed_data.append([date, exchange, commodity, price, month, year])

    # Convert the list into a DataFrame
    df = pd.DataFrame(parsed_data, columns=['Date', 'Exchange', 'Commodity', 'Price', 'Month', 'Year'])
    return df


def open_doc_and_parse_futures(path):
    try:
        doc = fitz.Document(path)
        extracted_data = extract_file_data(doc)
        return parse_to_table(extracted_data)
    except Exception as e:
        return None


def parse_directory(directory='pdfs'):
    errors = []
    dfs = []
    for filename in os.listdir(directory):
        try:
            path = os.path.join(directory, filename)
            df = open_doc_and_parse_futures(path)
            dfs.append(df)
        except Exception as e:
            errors.append(f"Error parsing {filename} error {e}")

    return dfs, errors

def final_battle():
    final, errors = parse_directory()
    final_table_data = pd.concat(final)


def get_table_download_link(df, sheet=None, filename=None):
    """Generates a link allowing the data in a given panda dataframe to be downloaded
    in:  dataframe
    out: href string
    """
    #writer = pd.ExcelWriter('unfiltered.xlsx', engine='xlsxwriter')
    #csv = df.to_excel(writer, sheet, filename)
    csv = df.to_csv()
    b64 = base64.b64encode(csv.encode()).decode()  # some strings <-> bytes conversions necessary here
    href = f'<a href="data:file/csv;base64,{b64}" download="result.csv">Download CSV File</a>'
    return href


# Create a title for the app
st.title('PDF Processing - Futures')

# Create a file uploader
uploaded_file = st.file_uploader("Choose a ZIP file", type="zip")

if uploaded_file is not None:
    # Extract the ZIP file
    with zipfile.ZipFile(uploaded_file, 'r') as zip_ref:
        zip_ref.extractall('pdfx')

    # Process the PDF files

    dfs, errors = parse_directory('pdfx')

    # Concatenate the dataframes
    final_table_data = pd.concat(dfs)

    # Get the unique commodities
    commodities = final_table_data['Commodity'].unique()
    
    commodities = np.append(commodities,"All")
    selected_commodity = st.selectbox('Select a commodity', commodities)
    if not selected_commodity or selected_commodity == "All":

        # Display the result in a table
        st.write(final_table_data)
        writer = pd.ExcelWriter('unfiltered.xlsx', engine='xlsxwriter')

        # Save the result to a CSV file
        final_table_data.to_csv('unfiltered-result.csv', index=False)

        # Add the download button
        st.markdown(get_table_download_link(final_table_data), unsafe_allow_html=True)
        st.success('File has been saved as unfiltered-result.csv')

    else:
        # filter
        dfx = final_table_data.copy()

        filtered_table = final_table_data[final_table_data['Commodity'] == selected_commodity]

        # Display the filtered table
        st.write(filtered_table)
        writer = pd.ExcelWriter('unfiltered.xlsx', engine='xlsxwriter')
        # Save the result to a CSV file

        # Add the download button
        st.markdown(get_table_download_link(filtered_table), unsafe_allow_html=True)
        st.success('File has been saved as filtered-result.csv')

    # Display any errors
    if errors:
        st.error('\n'.join(errors))


# Save the result to a CSV file
    final_table_data.to_csv('result.csv', index=False)
    st.success('File has been saved as result.csv')

    # Display any errors
    if errors:
        st.error('\n'.join(errors))
