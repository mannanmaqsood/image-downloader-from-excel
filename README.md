import pandas as pd
import requests
import os
from urllib.parse import urlparse

def download_images_from_xlsx(xlsx_filename, sheet_name, url_column, output_folder):
    """
    Download images from URLs in an Excel file and save them to a specified folder with real names.
    
    Parameters:
    xlsx_filename (str): Path to the Excel file.
    sheet_name (str): Name of the sheet containing URLs.
    url_column (str): Column name containing image URLs.
    output_folder (str): Folder to save downloaded images.
    """
    # Create output folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)
    
    # Read Excel file
    df = pd.read_excel(xlsx_filename, sheet_name=sheet_name)
    
    for url in df[url_column].dropna():
        url = url.strip()
        try:
            response = requests.get(url, timeout=10)
            if response.status_code == 200:
                # Extract filename from URL
                parsed_url = urlparse(url)
                filename = os.path.basename(parsed_url.path)
                
                # Save image with its real name
                image_path = os.path.join(output_folder, filename)
                with open(image_path, 'wb') as file:
                    file.write(response.content)
                print(f" Downloaded: {image_path}")
            else:
                print(f" Failed to download: {url}")
        except requests.RequestException:
            print(f" Error downloading: {url}")

# Example Usage
xlsx_file = 'file.xlsx'  # Replace with your Excel filename
sheet_name = 'Sheet1'  # Replace with your sheet name
url_column_name = 'Image_URL'  # Replace with your column name
output_directory = 'output_directory'  # Folder to save images

download_images_from_xlsx(xlsx_file, sheet_name, url_column_name, output_directory)
