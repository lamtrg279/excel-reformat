import pandas as pd

# Full output headers
headers = [
    "Line Type","customer","CSR","description","manufacturingLocation","promiseDate",
    "jobType","poNum","jobPart/@jobPart","jobpart/@qtyOrdered","jobShipment/@name",
    "jobshipment/@shipDate","jobShipment/@shipmentType","jobShipment/@trackingNumber",
    "jobShipment/@cost","jobShipment/@firstName","jobShipment/@lastName",
    "jobShipment/@address1","jobShipment/@address2","jobShipment/@address3",
    "jobShipment/@city","jobShipment/@state","jobShipment/@zip","jobShipment/@country",
    "phone","email","jobShipment/carton/@count","jobShipment/@shipVia", "jobMaterial/@inventoryItem",
    "jobMaterial/@plannedQuantity","ccJob","ccJobPart",
    "jobShipment/Carton/cartonContent@jobmaterial",
    "jobShipment/cartonContent/@quantity"
]

# Columns in output to fill
job_columns = ["description", "jobShipment/@address1", "jobShipment/@city", "jobShipment/@state","jobShipment/@zip", "poNum"]
input_file = pd.read_excel("input.xlsx")

all_row = []
try:
    for idx, row in input_file.iterrows():
        # ---- Process job row ----
        job_values = row.iloc[:6].tolist() # Get data from first 6 columns for mapping
        job_row = {col: "" for col in headers} # Job row for job data with empty values for all columns 

        job_row["jobShipment/@shipmentType"] = "8" 
        job_row["jobShipment/@country"] = "1"    
        job_row["jobShipment/carton/@count"] = "1"
        job_row["Line Type"] = "J"
        for job_col, job_val in zip(job_columns, job_values): # Map the 6 columns to the specified job columns in the output
            job_row[job_col] = job_val  # Fill the job row with values from the input 
        all_row.append(job_row)

        # ---- Process item rows ----
        item_values = row.index[6:] # Get item columns starting from the 7th column

        for item_col in item_values: # Loop through item columns and create item rows for each non-empty item column
            quantity = row[item_col] if item_col in row else None
            if pd.notna(quantity) and str(quantity).strip() not in ["", "nan", "NaN", "0"]:
                item_row = {col: "" for col in headers} # Initialize item row with empty values
                item_row["Line Type"] = "M"
                item_row["jobShipment/Carton/cartonContent@jobmaterial"] = item_col 
                item_row["jobShipment/cartonContent/@quantity"] = quantity 
                item_row["jobMaterial/@inventoryItem"] = item_col 
                item_row["jobMaterial/@plannedQuantity"] = quantity 
                all_row.append(item_row) # Add the item row to the output list

    df_out = pd.DataFrame(all_row, columns=headers)
    df_out.to_excel("output.xlsx", index=False)
    print("All rows processed successfully. Output saved to output.xlsx")
except Exception as e:
    print(f"Error processing row {idx}: {e}")

