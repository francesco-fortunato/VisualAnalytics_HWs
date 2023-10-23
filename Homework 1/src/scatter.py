import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os  

# Get the current working directory (directory where the script is located)
current_directory = os.path.dirname(os.path.abspath(__file__))

# Define the path to the Excel file relative to the current directory
excel_file_path = os.path.join(current_directory, '../dataframe/Exams.xlsx')

# Load the Excel file with multiple sheets (each person's scores)
excel_file = pd.ExcelFile(excel_file_path)

# List of sheet names (each sheet corresponds to a person's scores)
sheet_names = excel_file.sheet_names

# Create an empty list to store DataFrames for each sheet
dataframes = []

cumulative_weighted_means_data = []

# Iterate through each person's data
for sheet_index, sheet_name in enumerate(sheet_names):
    # Load the data from the sheet
    df = excel_file.parse(sheet_name)
    
    # Convert the 'Date' column to datetime
    df['Date'] = pd.to_datetime(df['Date'], format='%d/%m/%Y')
    
    # Sort the data by date for temporal distribution
    df = df.sort_values(by='Date')
    
    # Initialize cumulative variables
    cumulative_weighted_sum = 0
    cumulative_credits_sum = 0
    cumulative_weighted_means = []

    # Iterate through the rows and calculate cumulative weighted mean and standard deviation
    for index, row in df.iterrows():
        cumulative_weighted_sum += row['Grade'] * row['Credits']
        cumulative_credits_sum += row['Credits']
        cumulative_weighted_mean = cumulative_weighted_sum / cumulative_credits_sum
        cumulative_weighted_means.append(cumulative_weighted_mean)
    
    # Add the cumulative weighted mean to the DataFrame
    df['Cumulative Weighted Mean'] = cumulative_weighted_means

    # Print and display local performance
    print(f'{sheet_name} - Cumulative Weighted Mean (Last Entry): {cumulative_weighted_means[-1]:.2f}')
        
    # Calculate and store global performance data for this person
    average_grade = df['Grade'].mean()
    
    # global_performance_data.append({
    #     'Person': sheet_name,
    #     'Average Grade': average_grade
    # })

    # Store Cumulative Weighted Mean data for each person
    cumulative_weighted_means_data.append({
        'Person': sheet_name,
        'Cumulative Weighted Mean': cumulative_weighted_means[-1]
    })

# Iterate through each sheet (person's scores) and load it into a DataFrame
for sheet_name in sheet_names:
    df = excel_file.parse(sheet_name)
    df['Person'] = sheet_name  # Add a 'Person' column to store the person's name
    dataframes.append(df)

# Concatenate all DataFrames into one
combined_df = pd.concat(dataframes)

# Calculate the overall average weighted mean
overall_average_weighted_mean = sum(
    person_data['Cumulative Weighted Mean'] for person_data in cumulative_weighted_means_data
) / len(cumulative_weighted_means_data)

# Calculate the overall average weighted mean
overall_mean = combined_df['Grade'].mean()

# Calculate the count of exams for each person's performance
exams_count = combined_df.groupby(['Person', 'Grade']).size().reset_index(name='Count')

# Create a scatter plot comparing individual performance against group performance
plt.figure(figsize=(10, 6))
ax = sns.scatterplot(
    data=exams_count,
    x='Person',
    y='Grade',
    hue='Person',
    size='Count',
    sizes=(20, 300),  # Define the size range for markers
    palette='Set1',
    marker='o',
)
plt.axhline(overall_average_weighted_mean, color='gray', linestyle='-', label=f'Group Avg. Weighted Mean: {overall_average_weighted_mean:.2f}')

# Customize plot labels and title
ax.set_xlabel('')
ax.set_ylabel('Exam Scores')
ax.set_title('Individual vs. Group Performance Scatter Plot')

# Display the legend
# Rotate x-axis labels for better readability
plt.xticks(rotation=0)

# Set grid lines to be below the data points
ax.set_axisbelow(True)

plt.tight_layout(rect=[0, 0, 0.4, 1])
ax.legend(loc='upper right', bbox_to_anchor=(1.4, 1), fontsize=8)

# Specify the file path where you want to save the plot
output_file = os.path.join(current_directory,'../img/scatter_plot.png')

# Save the plot to the specified file
fig = plt.gcf()
fig.set_size_inches((20 , 11.25), forward=False) #1920x1080
fig.savefig(output_file, dpi=500, bbox_inches='tight')


# Show the plot
plt.show()
