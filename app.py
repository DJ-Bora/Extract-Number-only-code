import pandas as pd
import re

def extract_dag_point_from_excel(excel_file='Attendance.xlsx'):
    """
    Load Excel sheet and extract Dag and Point values from BOTH
    'Check-Out Status' and 'Check-Out Remark' columns.
    """
    try:
        df = pd.read_excel(excel_file).fillna("")
        print(f"Excel file '{excel_file}' loaded successfully with {len(df)} rows.")

        if 'Check-Out Status' not in df.columns:
            print("❌ Error: 'Check-Out Status' column not found.")
            return None
        if 'Check-Out Remark' not in df.columns:
            print("⚠️ Warning: 'Check-Out Remark' column not found. Using only Check-Out Status.")

        # More permissive regex patterns
        dag_regex = re.compile(
            r"(?:Dag|DAG|Daag|Daga|Dagg|Dage|Dags)[\:\-\_\s=\/\.]*(\d+)|"
            r"(\d+)\s+dag|Total\s+dag\s+(\d+)", re.IGNORECASE
        )

        # ✅ Improved point regex: catch multiple numbers & extra symbols
        point_regex = re.compile(
            r"(?:point|points?|ponit|poin|poit|pont|poits|colect|poins)[\:\-\_\s=\/\.]*(\d+)",
            re.IGNORECASE
        )

        dag_values, point_values = [], []

        for _, row in df.iterrows():
            status = row['Check-Out Status'].strip()
            remark = row.get('Check-Out Remark', '').strip()

            combined_text = f"{status} {remark}".strip()
            dag_val, point_val = None, None

            # --- Handle Nill / Zero cases
            if not combined_text:
                dag_values.append(None)
                point_values.append(None)
                continue
            if re.search(r'\bnill?\b', combined_text, re.IGNORECASE) or re.fullmatch(r'0+|00', combined_text):
                dag_values.append(0)
                point_values.append(0)
                continue

            # --- Handle "8/11" format
            if re.match(r'^\d+/\d+$', combined_text):
                d, p = map(int, combined_text.split('/'))
                dag_values.append(d)
                point_values.append(p)
                continue

            # --- Search Dag and Point in BOTH columns
            for text in (status, remark):
                if text:
                    if dag_val is None:
                        m = dag_regex.search(text)
                        if m:
                            dag_val = int(m.group(1) or m.group(2) or m.group(3))
                    # ✅ Capture ALL point matches, pick last one
                    point_matches = point_regex.findall(text)
                    if point_matches:
                        point_val = int(point_matches[-1])  # take last match

            dag_values.append(dag_val)
            point_values.append(point_val)

        df['Dag'] = dag_values
        df['Point'] = point_values

        output_file = excel_file.replace('.xlsx', '_updated.xlsx')
        df.to_excel(output_file, index=False)

        print(f"✅ Updated Excel saved as '{output_file}'")
        print(f"Total rows processed: {len(df)}")
        print(f"Dag Missing: {df['Dag'].isna().sum()} | Dag Zero: {(df['Dag']==0).sum()}")
        print(f"Point Missing: {df['Point'].isna().sum()} | Point Zero: {(df['Point']==0).sum()}")

        return df

    except Exception as e:
        print(f"❌ Error: {str(e)}")
        return None


def process_attendance():
    return extract_dag_point_from_excel('Attendance.xlsx')


if __name__ == "__main__":
    df = process_attendance()
    if df is not None:
        print("\nSample Output (first 10 rows):")
        print(df[['Check-Out Status','Check-Out Remark','Dag','Point']].head(10).to_string(index=False))
