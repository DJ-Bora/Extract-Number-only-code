import pandas as pd
import re

def extract_dag_point_plot_property(excel_file='Attendance.xlsx'):
    """
    Extract Dag, Point, Plot, and Property values from BOTH
    'Check-Out Status' and 'Check-Out Remark' columns.
    Missing values are left as blank.
    """
    try:
        df = pd.read_excel(excel_file).fillna("")
        print(f"Excel file '{excel_file}' loaded successfully with {len(df)} rows.")

        if 'Check-Out Status' not in df.columns:
            print("❌ Error: 'Check-Out Status' column not found.")
            return None
        if 'Check-Out Remark' not in df.columns:
            print("⚠️ Warning: 'Check-Out Remark' column not found. Using only Check-Out Status.")

        # --- Regex patterns
        dag_regex = re.compile(
            r"(?:Dag|DAG|Daag|Daga|Dagg|Dage|Dags)[\:\-\_\s=\/\.]*(\d+)|"
            r"(\d+)\s+dag|Total\s+dag\s+(\d+)", re.IGNORECASE
        )

        point_regex = re.compile(
            r"(?:point|points?|ponit|poin|poit|pont|poits|colect|poins)[\:\-\_\s=\/\.]*(\d+)",
            re.IGNORECASE
        )

        plot_regex = re.compile(
            r"(?:Plot|plot|Plt|pl|Plott|plt\.?|Plot#|Plt\-)[\:\-\_\s=\/\.]*(\d+)",
            re.IGNORECASE
        )

        property_regex = re.compile(
            r"(?:Property|property|Prop|prop|Prp|Prop#|Property\-)[\:\-\_\s=\/\.]*(\d+)",
            re.IGNORECASE
        )

        # --- Initialize lists
        dag_values, point_values, plot_values, property_values = [], [], [], []

        for _, row in df.iterrows():
            status = row['Check-Out Status'].strip()
            remark = row.get('Check-Out Remark', '').strip()
            combined_text = f"{status} {remark}".strip()

            dag_val = point_val = plot_val = property_val = ""

            # --- Handle empty / nill
            if not combined_text or re.search(r'\bnill?\b', combined_text, re.IGNORECASE) or re.fullmatch(r'0+|00', combined_text):
                dag_values.append("")
                point_values.append("")
                plot_values.append("")
                property_values.append("")
                continue

            # --- Handle "8/11" format
            if re.match(r'^\d+/\d+$', combined_text):
                d, p = map(int, combined_text.split('/'))
                dag_val, point_val = d, p

            # --- Search Dag, Point, Plot, Property in BOTH columns
            for text in (status, remark):
                if text:
                    if dag_val == "":
                        m = dag_regex.search(text)
                        if m:
                            dag_val = int(m.group(1) or m.group(2) or m.group(3))

                    point_matches = point_regex.findall(text)
                    if point_matches:
                        point_val = int(point_matches[-1])

                    if plot_val == "":
                        m = plot_regex.search(text)
                        if m:
                            plot_val = int(m.group(1))

                    if property_val == "":
                        m = property_regex.search(text)
                        if m:
                            property_val = int(m.group(1))

            dag_values.append(dag_val)
            point_values.append(point_val)
            plot_values.append(plot_val)
            property_values.append(property_val)

        # --- Add new columns
        df['Dag'] = dag_values
        df['Point'] = point_values
        df['Plot'] = plot_values
        df['Property'] = property_values

        output_file = excel_file.replace('.xlsx', '_updated.xlsx')
        df.to_excel(output_file, index=False)

        print(f"✅ Updated Excel saved as '{output_file}'")
        print(f"Total rows processed: {len(df)}")
        print(f"Dag Missing: {(df['Dag']=="").sum()}")
        print(f"Point Missing: {(df['Point']=="").sum()}")
        print(f"Plot Missing: {(df['Plot']=="").sum()}")
        print(f"Property Missing: {(df['Property']=="").sum()}")

        return df

    except Exception as e:
        print(f"❌ Error: {str(e)}")
        return None


def process_attendance():
    return extract_dag_point_plot_property('Attendance.xlsx')


if __name__ == "__main__":
    df = process_attendance()
    if df is not None:
        print("\nSample Output (first 10 rows):")
        print(df[['Check-Out Status','Check-Out Remark','Dag','Point','Plot','Property']].head(10).to_string(index=False))
