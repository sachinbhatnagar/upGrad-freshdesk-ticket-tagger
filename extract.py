import os
import sys
import pandas as pd
from extractPipedData import extract_piped_data


def extract_tickets(file_name, sheet_name, output_file):
    if not os.path.exists(file_name):
        print("File does not exist. Please check the filename and try again.")
        sys.exit()

    df = pd.read_excel(file_name, sheet_name=sheet_name, engine="openpyxl")
    extracted = []

    for _, row in df.iterrows():
        processed_data = extract_piped_data(str(row["Description"]))

        for data in processed_data:
            msg = data["message"].strip()
            if msg != "":
                extracted.append(
                    [
                        row.iloc[0],
                        row.iloc[1],
                        row.iloc[2],
                        row.iloc[3],
                        row.iloc[4],
                        data["message"],
                        str(data["timecode"]),
                        str(data["timestamp"]),
                        str(data["user"]),
                        row.iloc[6],
                        row.iloc[7],
                        row.iloc[8],
                        row.iloc[9],
                        row.iloc[10],
                        row.iloc[11],
                        row.iloc[12],
                        row.iloc[13],
                        row.iloc[14],
                    ]
                )

    df_extracted = pd.DataFrame(
        extracted,
        columns=[
            "ticketId",
            "workshopId",
            "subject",
            "sessionNumber",
            "emailId",
            "message",
            "timecode",
            "timestamp",
            "user",
            "status",
            "agent",
            "created",
            "workshopManager",
            "workshopDate",
            "workshopStatus",
            "upgradBuddy",
            "fullName",
            "contactEmail",
        ],
    )

    df_extracted.to_excel(output_file, index=False, engine="openpyxl")

    return len(extracted)
