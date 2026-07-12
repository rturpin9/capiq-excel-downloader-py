import pandas as pd
import numpy as np

df = pd.read_csv('mag7_comps.csv')

results = []
for _, row in df.iterrows():
    emp = row.get('Employees')
    ltm_rev = row.get('LTM Rev')
    ni = row.get('Net Income')
    company = row['Company']

    rev_per_emp = (ltm_rev * 1_000_000 / emp) if pd.notna(emp) and pd.notna(ltm_rev) and emp > 0 else None
    ni_per_emp = (ni * 1_000_000 / emp) if pd.notna(emp) and pd.notna(ni) and emp > 0 else None

    results.append({
        'Company': company,
        'Employees': int(emp) if pd.notna(emp) else None,
        'LTM Rev': ltm_rev,
        'Net Income': ni,
        'Rev/Employee': rev_per_emp,
        'Profit/Employee': ni_per_emp,
    })

# Sort by Rev/Employee descending
results.sort(key=lambda x: x['Rev/Employee'] or 0, reverse=True)

def fd(v):
    return f"${v:,.0f}" if v else "N/A"

def fc(v):
    return f"{v:,}" if v else "N/A"

print("| Company | Employees | LTM Revenue ($M) | Net Income ($M) | Rev / Employee | Profit / Employee |")
print("| --- | ---: | ---: | ---: | ---: | ---: |")
for r in results:
    print(f"| {r['Company']} | {fc(r['Employees'])} | {fd(r['LTM Rev'])} | {fd(r['Net Income'])} | {fd(r['Rev/Employee'])} | {fd(r['Profit/Employee'])} |")

emps = [r['Employees'] for r in results if r['Employees']]
revs = [r['Rev/Employee'] for r in results if r['Rev/Employee']]
profs = [r['Profit/Employee'] for r in results if r['Profit/Employee']]
print(f"| **Average** | {np.mean(emps):,.0f} | | | {fd(np.mean(revs))} | {fd(np.mean(profs))} |")
print(f"| **Median** | {np.median(emps):,.0f} | | | {fd(np.median(revs))} | {fd(np.median(profs))} |")
