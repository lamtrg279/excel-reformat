import pandas as pd

INPUT_FILE = "input.xlsx"
OUTPUT_FILE = "output.xlsx"

# Output columns required by JDF
OUTPUT_HEADERS = [
    "Line Type","customer","CSR","description","manufacturingLocation","promiseDate",
    "jobType","poNum","jobPart/@jobPart","jobpart/@qtyOrdered","jobShipment/@name",
    "jobshipment/@shipDate","jobShipment/@shipmentType","jobShipment/@trackingNumber",
    "jobShipment/@cost","jobShipment/@firstName","jobShipment/@lastName",
    "jobShipment/@address1","jobShipment/@address2","jobShipment/@address3",
    "jobShipment/@city","jobShipment/@state","jobShipment/@zip","jobShipment/@country",
    "phone","email","jobShipment/carton/@count","jobShipment/@shipVia",
    "jobMaterial/@inventoryItem","jobMaterial/@plannedQuantity","ccJob","ccJobPart",
    "jobShipment/Carton/cartonContent@jobmaterial","jobShipment/cartonContent/@quantity"
]

# Mapping of your input → job row fields
JOB_MAP = {
    "description": "Job Description",
    "jobShipment/@address1": "Address",
    "jobShipment/@city": "City",
    "jobShipment/@state": "State",
    "jobShipment/@zip": "Zip Code",
    "poNum": "PO#"
}

def empty_row():
    """Creates a blank output row."""
    return {c: "" for c in OUTPUT_HEADERS}


# ----------------------------
# MAIN PROCESSING
# ----------------------------

df = pd.read_excel(INPUT_FILE, engine="openpyxl")
job_input_cols = set(JOB_MAP.values()) # Columns from input that are mapped to job fields
item_columns = [c for c in df.columns if c not in job_input_cols] # Columns that are not mapped to job fields are treated as item columns

output = []

for _, row in df.iterrows():

    # ---- Job Row ----
    j = empty_row()
    j["Line Type"] = "J"
    j["jobShipment/@shipmentType"] = "8"
    j["jobShipment/@country"] = "1"
    j["jobShipment/carton/@count"] = "1"

    # Fill mapped job fields
    for out_col, in_col in JOB_MAP.items():
        j[out_col] = row[in_col]

    output.append(j)

    # ---- Item Rows ----
    for col in item_columns:
        qty = row[col]

        if pd.notna(qty) and str(qty).strip() not in ("", "0", "nan", "NaN"):
            m = empty_row()
            m["Line Type"] = "M"
            m["jobShipment/Carton/cartonContent@jobmaterial"] = col
            m["jobShipment/cartonContent/@quantity"] = qty
            m["jobMaterial/@inventoryItem"] = col
            m["jobMaterial/@plannedQuantity"] = qty
            output.append(m)

# Export
pd.DataFrame(output, columns=OUTPUT_HEADERS).to_excel(OUTPUT_FILE, index=False, engine="openpyxl")
print(f"Done. Saved to {OUTPUT_FILE}")
