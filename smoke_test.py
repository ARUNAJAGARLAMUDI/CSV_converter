#smoke_test.py
 
 
import os
from pathlib import Path
import pandas as pd
import util
 
 
def run_smoke_test():
    sample = [
        {
            'p_number': 'P-1001',
            'short_description': 'Upgrade database to v12',
            'description': 'We will upgrade the primary DB to version 12 to improve performance.',
            'affected_customers': 'Finance, HR',
            'state': 'In Progress',
            'completion_code': 'ONGOING'
        },
        {
            'p_number': 'P-1002',
            'short_description': 'Fix login bug',
            'description': 'Users occasionally fail to login due to session token expiration.',
            'affected_customers': 'All web users',
            'state': 'Open',
            'completion_code': 'NEW'
        }
    ]
 
    df = pd.DataFrame(sample)
 
    out_dir = Path('test_output')
    out_dir.mkdir(exist_ok=True)
 
    # Individual files
    for _, row in df.iterrows():
        buf = util.create_docx(row)
        filename = out_dir / f"Project_{row['p_number']}.docx"
        with open(filename, 'wb') as f:
            f.write(buf.getvalue())
 
    # Combined file
    combined = util.create_combined_docx(df)
    combined_path = out_dir / 'All_Project_Summaries.docx'
    with open(combined_path, 'wb') as f:
        f.write(combined.getvalue())
 
    print(f"Smoke test completed. Generated {len(df)} individual files and 1 combined file in {out_dir.resolve()}")
 
 
if __name__ == '__main__':
    run_smoke_test()