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

# Set a custom color palette with different colors for each person
custom_palette = sns.color_palette("Set1")

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

# Create a box plot using seaborn
plt.figure(figsize=(12, 6))
sns.boxplot(data=combined_df, x='Person', y='Grade', palette=custom_palette)


# Calculate the overall average weighted mean
overall_average_weighted_mean = sum(
    person_data['Cumulative Weighted Mean'] for person_data in cumulative_weighted_means_data
) / len(cumulative_weighted_means_data)

# Calculate the overall average weighted mean
overall_mean = combined_df.groupby('Person')['Grade'].mean().mean()

# Add a line representing the overall average weighted mean
plt.axhline(overall_average_weighted_mean, color='gray', linestyle='-', label=f'Overall Avg. Weighted Mean: {overall_average_weighted_mean:.2f}')

# Calculate and add the mean for each person as a cross marker
for person, color in zip(sheet_names, custom_palette):
    person_mean = combined_df[combined_df['Person'] == person]['Grade'].mean()
    plt.plot([sheet_names.index(person)], [person_mean], marker='x', markersize=8, color="black", label=f'{person} Mean: {person_mean:.2f}')

# Customize plot labels and title
plt.xlabel('')
plt.ylabel('Exam Scores')
plt.title('Box Plot of Exam Scores by Person')

# Rotate x-axis labels for better readability
plt.xticks(rotation=0)

plt.tight_layout(rect=[0, 0, 0.7, 1])

# Show the plot
plt.grid(True)
plt.legend(loc="upper left", fontsize=8)

fig = plt.gcf()
fig.set_size_inches((10 , 11.25), forward=False) #1920x1080
fig.savefig(os.path.join(current_directory, '../img/box_plot.png'), dpi=500, bbox_inches='tight')

plt.show()
