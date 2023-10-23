import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import matplotlib.dates as mdates
import numpy as np
import os 

# Define a custom color palette
custom_palette = sns.color_palette("Set1")

# Set the custom color palette
sns.set_palette(custom_palette) 

# Get the current working directory (directory where the script is located)
current_directory = os.path.dirname(os.path.abspath(__file__))

# Define the path to the Excel file relative to the current directory
excel_file_path = os.path.join(current_directory, '../dataframe/Exams.xlsx')

# Load the Excel file with multiple sheets (each person's scores)
excel_file = pd.ExcelFile(excel_file_path)

# Specify the sheet name you want to work with
sheet_name = 'Francesco'

# Load the data from the selected sheet
df = excel_file.parse(sheet_name)

# Convert the 'Date' column to datetime
df['Date'] = pd.to_datetime(df['Date'], format='%d/%m/%Y')

# Sort the data by date for temporal distribution
df = df.sort_values(by='Date')

# Initialize cumulative variables
cumulative_weighted_sum = 0
cumulative_credits_sum = 0
cumulative_weighted_means = []
cumulative_sum = 0
cumulative_means = []

# Create an array of x-coordinates for bars with spacing
x_values = range(len(df))
x_values = [x - 0.2 for x in x_values]  # Offset for bar spacing

# Iterate through the rows and calculate cumulative weighted mean and standard deviation
for x, row in zip(x_values, df.iterrows()):
    index, row_data = row
    cumulative_sum += row_data['Grade']
    cumulative_mean = cumulative_sum / (index+1)
    cumulative_weighted_sum += row_data['Grade'] * row_data['Credits']
    cumulative_credits_sum += row_data['Credits']
    cumulative_weighted_mean = cumulative_weighted_sum / cumulative_credits_sum
    cumulative_means.append(cumulative_mean)
    cumulative_weighted_means.append(cumulative_weighted_mean)

# Add the cumulative weighted mean and standard deviation to the DataFrame
df['Cumulative Weighted Mean'] = cumulative_weighted_means
df['Cumulative Mean'] = cumulative_means

# Create the figure and axes
fig, ax = plt.subplots(figsize=(10, 6))

# Plot the temporal distribution with bar plots and offset for spacing
ax.bar(x_values, df['Grade'], width=0.4, color=custom_palette[1])
ax.plot(x_values, df['Cumulative Weighted Mean'], linestyle='-', label=f'Cumulative Weighted Mean', color=custom_palette[3])
ax.plot(x_values, df['Cumulative Mean'], linestyle='-', label=f'Cumulative Simple Mean', color=custom_palette[4])
ax.axhline([cumulative_weighted_means[-1]], label=f'Final Weighted Mean', color='gray') 

# Here: add a new tick with the required value
yticks = [18, 20, 22, 24, cumulative_weighted_means[-1],  26, 28, 30,]
yticklabels = [*ax.get_yticklabels(), cumulative_weighted_means[-1]]
ax.set_yticks(yticks)


ax.set_title('Temporal Distribution of Exam Scores with Cumulative Means')
ax.set_xlabel('')
ax.set_ylabel('Grade')
ax.legend(loc='upper left')
ax.grid(True)
ax.set_xticks(x_values)
ax.set_xticklabels(df['Date'].dt.strftime('%b %Y'), rotation=45)

# Display local performance
print(f'Cumulative Weighted Mean (Last Entry): {cumulative_weighted_means[-1]:.2f}')

# Set the grid to be below the bars
ax.set_axisbelow(True)

# Calculate the mean and standard deviation
mean_grade = df['Grade'].mean()

plt.ylim(17, 31)

print(np.sqrt(np.cov(df['Grade'], aweights=df['Credits'])))

# Specify the file path where you want to save the plot
output_file = os.path.join(current_directory,'../img/local_performance_francesco.png')

# Save the plot to the specified file
fig = plt.gcf()
fig.set_size_inches((20 , 11.25), forward=False) #1920x1080
fig.savefig(output_file, dpi=500, bbox_inches='tight')
plt.show()