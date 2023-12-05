import os
import pandas as pd
import json
import numpy as np

def convert_to_json_serializable(value):
    if isinstance(value, np.int64):
        return int(value)
    elif pd.isna(value):
        return None
    else:
        return value

def generate_metadata(excel_file_path):
    try:
        # Read Excel file into a pandas DataFrame
        df = pd.read_excel(excel_file_path)

        # Create a dictionary to store assets grouped by instance_name
        assets_by_instance = {}

        # Iterate over rows in the DataFrame
        for index, row in df.iterrows():
            is_fed_by_names = [fed_by.strip() for fed_by in str(row['mqtt.physical_tag.asset.relationships.isFedBy']).split(',')]
            is_associated_with_name = str(row['mqtt.physical_tag.asset.relationships.isPartOf']).strip()

            # Determine the instance_name for grouping assets
            instance_name = row['mqtt.physical_tag.asset.instance_name']

            # Add asset to the corresponding instance in the dictionary
            if instance_name not in assets_by_instance:
                assets_by_instance[instance_name] = {'metadata': instance_name, 'version': row['mqtt.version'], 'timestamp': row['mqtt.timestamp'], 'assets': {}}

            # Create the main asset for the instance
            asset_id = row['mqtt.physical_tag.asset.instance_name']
            asset = {
                "instname": asset_id,
                "vendorname": convert_to_json_serializable(row['mqtt.physical_tag.asset.manufacturer']),
                "modelname": convert_to_json_serializable(row['mqtt.physical_tag.asset.model']),
                "firmware": convert_to_json_serializable(row['mqtt.physical_tag.asset.firmware_version']),
                "software_version": convert_to_json_serializable(row['mqtt.physical_tag.asset.software_version']),
                "serial_number": convert_to_json_serializable(row['mqtt.physical_tag.asset.serial_number']),
                "eng_unit_type": convert_to_json_serializable(row['mqtt.physical_tag.asset.eng_unit_type']),
                "eng_asset_tag": convert_to_json_serializable(row['mqtt.physical_tag.asset.eng_asset_tag']),
                "location": {
                    "x_coord": convert_to_json_serializable(row['mqtt.physical_tag.asset.location.x_coord']),
                    "y_coord": convert_to_json_serializable(row['mqtt.physical_tag.asset.location.y_coord'])
                },
                "relationships": {
                    "hasLocation": convert_to_json_serializable(row['mqtt.physical_tag.asset.relationships.hasLocation']),
                    "isAssociatedWith": convert_to_json_serializable(row['mqtt.physical_tag.asset.relationships.isAssociatedWith']),
                    "isPartOf": convert_to_json_serializable(row['mqtt.physical_tag.asset.relationships.isPartOf']),
                    "isFedBy": [convert_to_json_serializable(fed_by.strip()) for fed_by in is_fed_by_names]
                }
            }

            # Add the main asset to the assets dictionary for the instance
            assets_by_instance[instance_name]['assets'][asset_id] = asset

            # Create assets for isFedBy
            for is_fed_by_name in is_fed_by_names:
                if not df[df['mqtt.physical_tag.asset.instance_name'] == is_fed_by_name].empty:
                    fed_by_asset_row = df[df['mqtt.physical_tag.asset.instance_name'] == is_fed_by_name].iloc[0]

                    fed_by_asset_id = fed_by_asset_row['mqtt.physical_tag.asset.instance_name']

                    # Create fed-by asset dictionary
                    fed_by_asset = {
                        "instname": fed_by_asset_id,
                        "vendorname": convert_to_json_serializable(fed_by_asset_row['mqtt.physical_tag.asset.manufacturer']),
                        "modelname": convert_to_json_serializable(fed_by_asset_row['mqtt.physical_tag.asset.model']),
                        "firmware": convert_to_json_serializable(fed_by_asset_row['mqtt.physical_tag.asset.firmware_version']),
                        "software_version": convert_to_json_serializable(fed_by_asset_row['mqtt.physical_tag.asset.software_version']),
                        "serial_number": convert_to_json_serializable(fed_by_asset_row['mqtt.physical_tag.asset.serial_number']),
                        "eng_unit_type": convert_to_json_serializable(fed_by_asset_row['mqtt.physical_tag.asset.eng_unit_type']),
                        "eng_asset_tag": convert_to_json_serializable(fed_by_asset_row['mqtt.physical_tag.asset.eng_asset_tag']),
                        "location": {
                            "x_coord": convert_to_json_serializable(fed_by_asset_row['mqtt.physical_tag.asset.location.x_coord']),
                            "y_coord": convert_to_json_serializable(fed_by_asset_row['mqtt.physical_tag.asset.location.y_coord'])
                        },
                        "relationships": {
                            "hasLocation": convert_to_json_serializable(fed_by_asset_row['mqtt.physical_tag.asset.relationships.hasLocation']),
                            "isAssociatedWith": convert_to_json_serializable(fed_by_asset_row['mqtt.physical_tag.asset.relationships.isAssociatedWith']),
                            "isPartOf": convert_to_json_serializable(fed_by_asset_row['mqtt.physical_tag.asset.relationships.isPartOf']),
                            "isFedBy": None  # Leave as None for fed-by assets
                        }
                    }

                    # Add the fed-by asset to the assets dictionary for the instance
                    assets_by_instance[instance_name]['assets'][fed_by_asset_id] = fed_by_asset

            # Create asset for isAssociatedWith
            if not df[df['mqtt.physical_tag.asset.instance_name'] == is_associated_with_name].empty:
                associated_with_asset_row = df[df['mqtt.physical_tag.asset.instance_name'] == is_associated_with_name].iloc[0]

                associated_with_asset_id = associated_with_asset_row['mqtt.physical_tag.asset.instance_name']

                # Create associated with asset dictionary
                associated_with_asset = {
                    "instname": associated_with_asset_id,
                    "vendorname": convert_to_json_serializable(associated_with_asset_row['mqtt.physical_tag.asset.manufacturer']),
                    "modelname": convert_to_json_serializable(associated_with_asset_row['mqtt.physical_tag.asset.model']),
                    "firmware": convert_to_json_serializable(associated_with_asset_row['mqtt.physical_tag.asset.firmware_version']),
                    "software_version": convert_to_json_serializable(associated_with_asset_row['mqtt.physical_tag.asset.software_version']),
                    "serial_number": convert_to_json_serializable(associated_with_asset_row['mqtt.physical_tag.asset.serial_number']),
                    "eng_unit_type": convert_to_json_serializable(associated_with_asset_row['mqtt.physical_tag.asset.eng_unit_type']),
                    "eng_asset_tag": convert_to_json_serializable(associated_with_asset_row['mqtt.physical_tag.asset.eng_asset_tag']),
                    "location": {
                        "x_coord": convert_to_json_serializable(associated_with_asset_row['mqtt.physical_tag.asset.location.x_coord']),
                        "y_coord": convert_to_json_serializable(associated_with_asset_row['mqtt.physical_tag.asset.location.y_coord'])
                    },
                    "relationships": {
                        "hasLocation": convert_to_json_serializable(associated_with_asset_row['mqtt.physical_tag.asset.relationships.hasLocation']),
                        "isAssociatedWith": convert_to_json_serializable(associated_with_asset_row['mqtt.physical_tag.asset.relationships.isAssociatedWith']),
                        "isPartOf": convert_to_json_serializable(associated_with_asset_row['mqtt.physical_tag.asset.relationships.isPartOf']),
                        "isFedBy": None  # Leave as None for associated with asset
                    }
                }

                # Add the associated with asset to the assets dictionary for the instance
                assets_by_instance[instance_name]['assets'][associated_with_asset_id] = associated_with_asset

        # Iterate over instances and create folders with metadata files
        for instance_name, metadata_data in assets_by_instance.items():
            metadata_json = json.dumps(metadata_data, default=convert_to_json_serializable, indent=2)

            # Construct the path for the folder and metadata file
            excel_directory = os.path.dirname(excel_file_path)
            asset_folder_path = os.path.join(excel_directory, instance_name)
            metadata_file_path = os.path.join(asset_folder_path, 'metadata.json')

            # Create the folder for the asset if it doesn't exist
            if not os.path.exists(asset_folder_path):
                os.makedirs(asset_folder_path)

            # Write JSON to the metadata file inside the folder
            with open(metadata_file_path, 'w') as file:
                file.write(metadata_json)

            print(f"Metadata file successfully created: {metadata_file_path}")

    except Exception as e:
        print(f"Error: {e}")

# Replace 'your_excel_file.xlsx' with the path to your Excel file
generate_metadata(r'C:\Users\daniel.glazier\OneDrive - ExcelRedstone\Documents\3ssGPT\site_model_data.xlsx')
