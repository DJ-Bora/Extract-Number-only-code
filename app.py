import pandas as pd
import re

def extract_dag_point_plot_property(excel_file='Attendance.xlsx'):
    """
    Extract Dag, Point, Plot, and Property values from BOTH
    'Check-Out Status' and 'Check-Out Remark' columns.
    Multiple numbers are captured as comma-separated strings.
    Leading zeros are removed (e.g., '00' -> '0', '02' -> '2').
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

        # Regex patterns
        dag_regex = re.compile(r"(?:Dag|DAG|Daag|Daga|Dagg|Dage|Dags)[\:\-\_\s=\/\.]*(\d+)|(\d+)\s+dag|Total\s+dag\s+(\d+)", re.IGNORECASE)
        point_regex = re.compile(r"(?:point|points?|ponit|poin|poit|pont|poits|colect|poins)[\:\-\_\s=\/\.]*(\d+)", re.IGNORECASE)
        plot_regex = re.compile(r"(?:Plot|plot|Plt|pl|Plott|Plots|plt\.?|Plot#|Plt\-)[\:\-\_\s=\/\.]*(\d+)", re.IGNORECASE)
        property_regex = re.compile(r"(?:Property|property|Prop|Properties|prop|Prp|Prop#|Property\-)[\:\-\_\s=\/\.]*(\d+)", re.IGNORECASE)

        def normalize_number(val):
            if val == "" or val is None:
                return None  # Use None for missing values
            try:
                return int(val)  # converts "00" -> 0, "02" -> 2
            except:
                return None

        # Initialize lists
        dag_values, point_values, plot_values, property_values = [], [], [], []

        for _, row in df.iterrows():
            status = str(row['Check-Out Status']).strip()
            remark = str(row.get('Check-Out Remark', '')).strip()
            combined_text = f"{status} {remark}".strip()

            dag_val = point_val = plot_val = property_val = None

            if not combined_text or re.search(r'\bnill?\b', combined_text, re.IGNORECASE) or re.fullmatch(r'0+|00', combined_text):
                dag_values.append(None)
                point_values.append(None)
                plot_values.append(None)
                property_values.append(None)
                continue

            # Handle "8/11" format (Dag/Point)
            if re.match(r'^\d+/\d+$', combined_text):
                d, p = map(int, combined_text.split('/'))
                dag_val = d
                point_val = p

            # Extract from both columns
            for text in (status, remark):
                if text:
                    if dag_val is None:
                        dag_matches = dag_regex.findall(text)
                        if dag_matches:
                            for g1, g2, g3 in dag_matches:
                                dag_val = g1 or g2 or g3
                                if dag_val:
                                    dag_val = int(dag_val)
                                    break

                    if point_val is None:
                        point_matches = point_regex.findall(text)
                        if point_matches:
                            point_val = int(point_matches[-1])

                    if plot_val is None:
                        plot_matches = plot_regex.findall(text)
                        if plot_matches:
                            plot_val = ",".join([str(int(x)) for x in plot_matches])

                    if property_val is None:
                        property_matches = property_regex.findall(text)
                        if property_matches:
                            property_val = ",".join([str(int(x)) for x in property_matches])

            dag_values.append(dag_val)
            point_values.append(point_val)
            plot_values.append(plot_val)
            property_values.append(property_val)

        # Add columns
        df['Dag'] = dag_values
        df['Point'] = point_values
        df['Plot'] = plot_values
        df['Property'] = property_values

        # Convert to numeric where possible
        df['Dag'] = pd.to_numeric(df['Dag'], errors='coerce')
        df['Point'] = pd.to_numeric(df['Point'], errors='coerce')
        df['Plot'] = pd.to_numeric(df['Plot'], errors='coerce')
        df['Property'] = pd.to_numeric(df['Property'], errors='coerce')

        output_file = excel_file.replace('.xlsx', '_updated.xlsx')
        df.to_excel(output_file, index=False)

        print(f"✅ Updated Excel saved as '{output_file}'")
        print(f"Total rows processed: {len(df)}")
        print(f"Dag Missing: {(df['Dag'].isna()).sum()}")
        print(f"Point Missing: {(df['Point'].isna()).sum()}")
        print(f"Plot Missing: {(df['Plot'].isna()).sum()}")
        print(f"Property Missing: {(df['Property'].isna()).sum()}")

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
