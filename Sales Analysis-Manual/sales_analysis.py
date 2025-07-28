# Imported Useful Library's
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# Importing Datadet
data = pd.read_excel("SuperStore Sales DataSet.xlsx")
df = pd.DataFrame(data)

# Data Cleaning (Removing Duplicates and Null Values)
df.drop_duplicates(subset = "Customer ID", inplace = True)
df.dropna(axis = 1, inplace = True)

#1 total sales and profit for each product category
grouped1 = df.groupby("Category")[["Sales", "Profit"]].sum().reset_index()
melted = grouped1.melt(id_vars = "Category", value_vars = ["Sales","Profit"], var_name = "SP", value_name = "Amount")

sns.barplot(data = melted, x = "Category", y = "Amount", hue="SP", palette=["mediumpurple", "dodgerblue"])
plt.title("Total Sales vs Total Profit by Category", fontsize = 9)
plt.xlabel("Category")
plt.ylabel("Amount (USD)")
plt.show()

#2 total sales by region
colors = ["lightblue","Orange"]
grouped2 = df.groupby("Region")["Sales"].sum()
grouped2 = grouped2.sort_values(ascending = False)
avg_sales = grouped2.values.mean()

plt.bar(grouped2.index, grouped2.values, color = colors)
plt.axhline(avg_sales, color='gray', linestyle='--', label='Avg Sales', linewidth=2) # For Creating Horizontal line
plt.xlabel("Region")
plt.title("Sales By Region With Below Average Highlight",fontsize = 9)
plt.ylabel("Total Sales")
plt.show()

#3 sales distribution across different customer segments
grouped3 = df.groupby("Segment")[["Sales","Profit"]].sum().reset_index()
melted1 = grouped3.melt(id_vars = "Segment", value_vars = ["Sales","Profit"], var_name = "SG_name", value_name = "Amount")

sns.barplot(data = melted1, x = "Segment", y = "Amount", hue="SG_name", palette=["mediumpurple", "dodgerblue"])
plt.title("Sales Across Diff Segment By Sales/Profit", fontsize = 9)
plt.show()

#4 average profit per order for each shipping mode
grouped4 = df.groupby("Ship Mode")[["Sales","Profit"]].sum().reset_index()
melted2 = grouped4.melt(id_vars = "Ship Mode", value_vars = ["Sales","Profit"], var_name = "SM", value_name = "Value")

sns.barplot(data = melted2, x = "Ship Mode", y = "Value", hue="SM", palette=["mediumpurple", "dodgerblue"])
plt.title("Sales Across Diff Ship-Mode By Sales/Profit", fontsize = 9)
plt.show()

#5 Correlation Coefficient between sales and profit
sns.scatterplot(data = df, x = df["Sales"], y = df["Profit"])
plt.title("Relationship b/w Sales and Profit", fontsize = 9)
crr = df["Sales"].corr(df["Profit"]) # Correlation b/w Sales and Profit
plt.show()

#6 the top 10 cities by sales
grouped5 = df.groupby("City")["Sales"].sum().sort_values(ascending = False).head(5)
plt.bar(grouped5.index, grouped5.values)
plt.xticks(rotation = 45)
plt.show()

#7 Distribution of sales by payment mode
grouped6 = df.groupby("Payment Mode")["Sales"].sum()
explode = [0.03, 0, 0]
plt.pie(grouped6.values, labels = grouped6.index, autopct = "%.2f", explode = explode)
plt.title("Total Sales By Payment Mode")
plt.show()









