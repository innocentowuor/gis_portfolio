# Mountain Goat AGOL Layer Download Script (Simplified)
# Author: Innocent Owuor
# Date: 27th March 2025
# Updated: 
# Description:
#   Downloads ArcGIS Feature layers into a flat structure in a geodatabase (back_ups.gdb)

# %% -------------------------------------------------------------------------------------
# Imports, Credentials, Configurations
# ----------------------------------------------------------------------------------------
import arcpy 
from arcgis.gis import GIS
from arcgis.features import FeatureLayer
import os
from json import load

# Load Credentials from a credentials JSON file 
with open(r"path to JSON file") as credentials_file:
    credentials = load(credentials_file)

agol_username, agol_password = credentials["agol_username"], credentials["agol_password"]
url = "your arcgis online URL"

# Paths Configuration
class Paths:
    data_folder = "Path to the directory where the geodatabase will be saved"   
    workgdb = f"{data_folder}\\back_ups.gdb"  # Target geodatabase

# AGOL Layer URLs
agol_layers = {
    "Feature Layer Name": "Feature Layer URL, Usually looks like https://services6.arcgis.com/ubm4tcTYICKBpist/arcgis/rest/services/.... ",
    
}

# %% -------------------------------------------------------------------------------------
# Functions
# ----------------------------------------------------------------------------------------

def connect_to_agol(url, username, password):
    """This function connects to ArcGIS Online using a user's credentials"""
    try:
        gis = GIS(url, username, password, verify_cert=False)
        print("Connected to ArcGIS Online.")
        return gis
    except Exception as e:
        print(f"Failed to connect to AGOL: {e}")
        return None

def clear_geodatabase(gdb_path):
    """Delete everything in the geodatabase that was used previously for backups."""
    try:
        arcpy.env.workspace = gdb_path
        for fc in arcpy.ListFeatureClasses():
            arcpy.Delete_management(fc)
            print(f"Deleted feature class: {fc}")
        print("Geodatabase cleared.")
    except Exception as e:
        print(f"Error clearing geodatabase: {e}")

def download_agol_layer(layer_name, layer_url, output_gdb):
    """Download AGOL feature layer and save to GDB root."""
    try:
        if "tiles.arcgis.com" in layer_url:
            print(f"Skipping raster layer: {layer_name}")
            return

        fl = FeatureLayer(layer_url)
        print(f"Accessing: {layer_name}")

        features = fl.query(where="1=1", out_fields="*", return_geometry=True)
        print(f"Downloaded {len(features)} features.")

        output_fc = os.path.join(output_gdb, layer_name)
        if arcpy.Exists(output_fc):
            arcpy.Delete_management(output_fc)

        features.save(output_gdb, layer_name)
        print(f"Saved to: {output_fc}")

    except Exception as e:
        print(f"Error downloading {layer_name}: {e}")

# %% -------------------------------------------------------------------------------------
# Main Execution
# ----------------------------------------------------------------------------------------

if __name__ == "__main__":
    arcpy.env.workspace = Paths.workgdb
    arcpy.env.overwriteOutput = True

    clear_geodatabase(Paths.workgdb)

    gis = connect_to_agol(url, agol_username, agol_password)
    if not gis:
        print("Exiting: failed AGOL connection.")
        exit(1)

    for layer_name, layer_url in agol_layers.items():
        print(f"Processing: {layer_name}")
        download_agol_layer(layer_name, layer_url, Paths.workgdb)

    print("AGOL Layer download complete.")

# %%
