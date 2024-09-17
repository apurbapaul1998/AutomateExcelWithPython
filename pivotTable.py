import pandas as pd
df=pd.read_csv(r"C:\Users\apurb\Downloads\superMarket-sales\supermarket_sales.csv")



df=df[['Gender', 'Product line', 'Total']]
# print(df)

pivot_table=df.pivot_table(index="Gender", columns='Product line', values='Total', aggfunc='sum')
# print(pivot_table)

pivot_table.to_excel('Pivot_table.xlsx','Pivot_report', startrow=4)