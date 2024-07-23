import os
import subprocess
import openpyxl
from openpyxl import Workbook

def get_image_metadata_jpg(image_path):
    # Run exiftool to get the metadata
    result = subprocess.run(['exiftool', image_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    
    if result.returncode != 0:
        print(f"Error reading metadata for {image_path}: {result.stderr}")
        return None

    metadata = {}
    for line in result.stdout.splitlines():
        if 'Bits Per Sample' in line:
            metadata['Bits Per Pixel'] = line.split(':')[-1].strip()
        elif 'Sub Sampling' in line:
            metadata['Subsampling'] = line.split(': ')[-1].strip()
        elif 'Image Size' in line:
            metadata['Image Dimensions'] = line.split(':')[-1].strip()
        elif 'ICC' in line:
            metadata['Color Profile'] = line.split(':')[-1].strip()
    
    return metadata

def get_image_metadata_png(image_path):
    # Run exiftool to get the metadata
    result = subprocess.run(['exiftool', image_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    
    if result.returncode != 0:
        print(f"Error reading metadata for {image_path}: {result.stderr}")
        return None

    metadata = {}
    for line in result.stdout.splitlines():
        if 'Bit Depth' in line:
            metadata['Bits Per Pixel'] = line.split(':')[-1].strip()
        elif 'Sub Sampling' in line:
            metadata['Subsampling'] = line.split(':')[-1].strip()
        elif 'Image Size' in line:
            metadata['Image Dimensions'] = line.split(':')[-1].strip()
        elif 'Color Type' in line:
            metadata['Color Profile'] = line.split(':')[-1].strip()
    
    return metadata

def get_image_metadata_avif(image_path):
    # Run exiftool to get the metadata
    result = subprocess.run(['exiftool', image_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    
    if result.returncode != 0:
        print(f"Error reading metadata for {image_path}: {result.stderr}")
        return None

    metadata = {}
    for line in result.stdout.splitlines():
        if 'Image Pixel Depth' in line:
            metadata['Bits Per Pixel'] = line.split(':')[-1].strip()
        elif 'Chroma Format' in line:
            metadata['Subsampling'] = line.split(': ')[-1].strip()
        elif 'Image Size' in line:
            metadata['Image Dimensions'] = line.split(':')[-1].strip()
        elif 'Color Profiles' in line:
            metadata['Color Profile'] = line.split(':')[-1].strip()
    
    return metadata

def write_to_excel(data, output_file):
    # Create a new Excel workbook and select the active worksheet
    wb = Workbook()
    ws = wb.active
    ws.title = "Image Metadata"

    # Define the header
    headers = ["Filename", "Bits Per Pixel", "Subsampling", "Image Dimensions", "Color Profile"]
    ws.append(headers)

    # Write the data
    for filename, metadata in data.items():
        row = [filename]
        row.append(metadata.get("Bits Per Pixel", "N/A"))
        row.append(metadata.get("Subsampling", "N/A"))
        row.append(metadata.get("Image Dimensions", "N/A"))
        row.append(metadata.get("Color Profile", "N/A"))
        ws.append(row)

    # Save the workbook
    wb.save(output_file)

def main(folder_path, output_file):
    image_metadata = {}
    
    # Iterate over the files in the folder
    for filename in os.listdir(folder_path):
        if filename.lower().endswith(('.png', '.jpg', '.jpeg', '.tiff', '.bmp', '.gif', '.avif')):
            image_path = os.path.join(folder_path, filename)
            if filename.lower().endswith(('.jpg')):
                metadata = get_image_metadata_jpg(image_path)
            elif filename.lower().endswith(('.png')):
                metadata = get_image_metadata_png(image_path)
            elif filename.lower().endswith(('.avif')):
                metadata = get_image_metadata_avif(image_path)
            if metadata:
                image_metadata[filename] = metadata

    # Write the metadata to the Excel file
    write_to_excel(image_metadata, output_file)
    print(f"Metadata has been written to {output_file}")

if __name__ == "__main__":
    # folder_path = "archive_cloudinary"          # Change this to the path of your image folder
    # output_file = "image_metadata_cloudinary.xlsx"      # Change this to your desired output Excel file name
    # folder_path = "dataset_classification"          # Change this to the path of your image folder
    # output_file = "image_metadata_before.xlsx"      # Change this to your desired output Excel file name
    folder_path = "archive_ShortPixelOptimized_glossy"          # Change this to the path of your image folder
    output_file = "image_metadata_shortpixel_glossy.xlsx"      # Change this to your desired output Excel file name
    main(folder_path, output_file)
