Asset Metadata Generator
Overview
The Asset Metadata Generator is a Python script designed to read data from an Excel spreadsheet containing information about physical assets and generate structured JSON metadata files for each asset. The metadata includes details such as asset properties, relationships, and location.

Prerequisites
Python 3.x
Pandas library (pip install pandas)
NumPy library (pip install numpy)
Installation
Clone or download the repository to your local machine.
bash
Copy code
git clone https://github.com/your-username/asset-metadata-generator.git
cd asset-metadata-generator
Install the required Python libraries.
bash
Copy code
pip install pandas numpy
Usage
Prepare your Excel spreadsheet with asset information. Ensure that the spreadsheet includes necessary columns such as mqtt.physical_tag.asset.instance_name, mqtt.physical_tag.asset.manufacturer, etc.

Modify the script parameters:

Replace 'your_excel_file.xlsx' in the script with the path to your Excel file.
Customize any other parameters based on your requirements.
Run the script.

bash
Copy code
python asset_metadata_generator.py
The script will create a folder for each asset, containing a metadata.json file with detailed asset information.
Customization
You can further customize the script to adapt to specific column names, handle different data types, or add additional features based on your needs.
License
This project is licensed under the MIT License - see the LICENSE file for details.
