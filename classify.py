import pandas as pd
import joblib
import os


def classify_tickets(file_name, model_pick, output_file):
    current_path = os.path.dirname(os.path.abspath(__file__))
    ds_model = os.path.join(current_path, "ds_model_v2.pkl")
    fsd_model = os.path.join(current_path, "fsd_model_v2.pkl")
    if model_pick == "Data Science":
        text_clf = joblib.load(ds_model)
    else:
        text_clf = joblib.load(fsd_model)

    df_new = pd.read_excel(file_name, engine="openpyxl")
    df_new.dropna(subset=["message"], inplace=True)
    df_new["category"] = text_clf.predict(df_new["message"])
    df_new.to_excel(output_file, index=False)
