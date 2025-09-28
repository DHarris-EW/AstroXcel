import numpy as np

def clean_dataframe(df, keep_columns):
    # Creates one dataframe for IT based on the cost excel file uploaded by the user
    
    # Replace 0.0, "", and " " with NaN 
    df = df.replace({0.0: np.nan, "": np.nan, " ": np.nan}, regex=False)
    
    # Keep only the specified columns and drop completely empty columns
    df = df[[col for col in keep_columns if col in df.columns]]
    # Drop rows where all cells in the row are NaN. Retain rows with actual data
    df = df.dropna(how='all')
    
    # Fill NaN values in rows with data that used to be 0.0, "", or " "
    df = df.fillna(0.00)
    
    return df