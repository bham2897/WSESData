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


# # Get the count of opportunities by region
# reporting_status_counts = data['Reporting Status'].value_counts()

# # Plot the count of opportunities by reporting status in a bar chart
# reporting_status_counts.plot(kind='bar', color=['red','lightblue'])
# plt.title('Count of Opportunities by Reporting Status')
# plt.xlabel('Reporting Status')
# plt.ylabel('Number of Opportunities')
# plt.show()

# # Plot the distribution of opportunities by region in a bar chart
# region_counts.plot(kind='bar', color='lightblue')
# plt.title('Distribution of Opportunities by Region')
# plt.xlabel('Region')
# plt.ylabel('Number of Opportunities')
# plt.show()

# Alternatively, you can use a treemap for a hierarchical view of the data
# Uncomment the following lines to use a treemap
# import squarify
# squarify.plot(sizes=region_counts, label=region_counts.index, color=sns.color_palette("Set3", len(region_counts)))
# plt.title('Distribution of Opportunities by Region')
# plt.show()


# # Plot the distribution of profit percentages in a box plot
# plt.boxplot(data['Profit %'], vert=False, widths=0.7, patch_artist=True, boxprops=dict(facecolor='lightgreen'))
# plt.title('Profit Percentage Distribution')
# plt.xlabel('Profit Percentage')
# plt.show()



# # Plot a scatter plot to visualize the correlation between sales value and profit
# plt.scatter(data['Sales Value in Million'], data['Profit'])
# plt.title('Correlation Between Sales Value and Profit')
# plt.xlabel('Sales Value in Million')
# plt.ylabel('Profit')
# plt.show()
            

# Get the mean relative strength in the segment for each industry
industry_strength = data.groupby('Industry')['Relative Strength in the segment'].mean()

# # Plot the relative strength in the segment across industries in a horizontal bar chart
# industry_strength.sort_values().plot(kind='barh', color='lightblue')
# plt.title('Relative Strength in the Segment Across Industries')
# plt.xlabel('Mean Relative Strength')
# plt.ylabel('Industry')
# plt.show()



# # Create a stacked bar chart for 'Leads Conversion Class' and 'WSES Proportion in Joint Bid'
# conversion_proportion = data.groupby(['Leads Conversion Class', 'WSES Proportion in Joint Bid']).size().unstack()

# # Plot the stacked bar chart
# conversion_proportion.plot(kind='bar', stacked=True, colormap='viridis')
# plt.title('Leads Conversion Class and WSES Proportion')
# plt.xlabel('Leads Conversion Class')
# plt.ylabel('Count')
# plt.show()



# # Assuming you have a date column (e.g., 'Date') in your dataset
# # Replace 'Date' with the actual date column name in your dataset

# # Convert 'Date' column to datetime format
# data['Date'] = pd.to_datetime(data['Date'])

# # Create a time series plot for opportunities over time
# opportunities_over_time = data.groupby(data['Date'].dt.to_period('M')).size()

# # Plot the time series plot
# opportunities_over_time.plot(kind='line', marker='o', color='purple')
# plt.title('Time Series Analysis for Opportunities Over Time')
# plt.xlabel('Time (Monthly)')
# plt.ylabel('Number of Opportunities')
# plt.show()



# # Get the mean profit for each product
# product_profit = data.groupby('Product')['Profit of Customer in Million'].mean()

# # Plot the comparison of profit across products in a bar chart
# product_profit.sort_values().plot(kind='bar', color='salmon')
# plt.title('Comparison of Profit Across Products')
# plt.xlabel('Product')
# plt.ylabel('Mean Profit (Million)')
# plt.show()



# # Assuming you have a hierarchical structure with 'Industry' and 'Product' columns
# # Replace 'Industry' and 'Product' with the actual column names in your dataset

# # Create a treemap for industry proportions
# industry_proportions = data.groupby(['Industry', 'Product']).size().reset_index(name='Count')

# # Use Seaborn's color palette
# colors = sns.color_palette('viridis')

# # Plot the treemap
# squarify.plot(sizes=industry_proportions['Count'], label=industry_proportions['Industry'] + ': ' + industry_proportions['Product'], color=colors)
# plt.title('Treemap for Industry Proportions')
# plt.show()