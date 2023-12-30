import pandas as pd
import matplotlib.pyplot as plt
import squarify
import seaborn as sns

# Read the data from the Excel file
data = pd.read_excel('/Users/divya/Downloads/WSES data.xlsx', sheet_name='Sheet1')

# Get the count of opportunities by sales outcome
sales_outcome_counts = data['Sales Outcome'].value_counts()

# Plot the distribution of sales outcomes in a pie chart
sales_outcome_counts.plot(kind='pie', autopct='%1.1f%%', startangle=90, explode=(0.1, 0), shadow=True)
plt.title('Distribution of Sales Outcomes')
plt.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
plt.show()


# Get the count of opportunities by region
reporting_status_counts = data['Reporting Status'].value_counts()

# Plot the count of opportunities by reporting status in a bar chart
reporting_status_counts.plot(kind='bar', color=['red','lightblue'])
plt.title('Count of Opportunities by Reporting Status')
plt.xlabel('Reporting Status')
plt.ylabel('Number of Opportunities')
plt.show()


# Get the count of opportunities by region
region_counts = data['Region'].value_counts()

# Plot the distribution of opportunities by region in a bar chart
region_counts.plot(kind='bar', color='lightblue')
plt.title('Distribution of Opportunities by Region')
plt.xlabel('Region')
plt.ylabel('Number of Opportunities')
plt.show()


# Plot the distribution of profit percentages in a box plot
plt.boxplot(data['Profit %'], vert=False, widths=0.7, patch_artist=True, boxprops=dict(facecolor='lightgreen'))
plt.title('Profit Percentage Distribution')
plt.xlabel('Profit Percentage')
plt.show()



# Plot a scatter plot to visualize the correlation between sales value and profit
plt.scatter(data['Sales Value in Million'], data['Profit %'])
plt.title('Correlation Between Sales Value and Profit')
plt.xlabel('Sales Value in Million')
plt.ylabel('Profit')
plt.show()
            

# Get the mean relative strength in the segment for each industry
industry_strength = data.groupby('Industry')['Relative Strength in the segment'].mean()

# Plot the relative strength in the segment across industries in a horizontal bar chart
industry_strength.sort_values().plot(kind='barh', color='lightblue')
plt.title('Relative Strength in the Segment Across Industries')
plt.xlabel('Mean Relative Strength')
plt.ylabel('Industry')
plt.show()


# Get the mean profit for each product
product_profit = data.groupby('Product')['Profit of Customer in Million'].mean()

# Plot the comparison of profit across products in a bar chart
product_profit.sort_values().plot(kind='bar', color='salmon')
plt.title('Comparison of Profit Across Products')
plt.xlabel('Product')
plt.ylabel('Mean Profit (Million)')
plt.show()
