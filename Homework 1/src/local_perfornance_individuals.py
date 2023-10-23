import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import matplotlib.dates as mdates  
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

# List of sheet names
sheet_names = excel_file.sheet_names

# Create a list to store dataframes for each person
dataframes = []

# Initialize lists to store the global performance data
global_performance_data = []

# Define a list of colors for Cumulative Weighted Mean lines
bar_colors = custom_palette[:len(sheet_names)]
line_colors = custom_palette[len(sheet_names):]

# Create a list to store Cumulative Weighted Mean data for each person
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
    
    # Add the cumulative weighted mean and standard deviation to the DataFrame
    df['Cumulative Weighted Mean'] = cumulative_weighted_means
    
    # Print and display local performance
    print(f'{sheet_name} - Cumulative Weighted Mean (Last Entry): {cumulative_weighted_means[-1]:.2f}')
    
    dataframes.append(df)
    
    # Calculate and store global performance data for this person
    average_grade = df['Grade'].mean()

    global_performance_data.append({
        'Person': sheet_name,
        'Average Grade': average_grade
    })

    # Store Cumulative Weighted Mean data for each person
    cumulative_weighted_means_data.append({
        'Person': sheet_name,
        'Cumulative Weighted Mean': cumulative_weighted_means
    })

# Concatenate the dataframes and reset the index
global_df = pd.concat(dataframes, keys=sheet_names)
global_df = global_df.reset_index()

# Convert global performance data to a DataFrame
global_performance_df = pd.DataFrame(global_performance_data)

# Create a single graph to display local performance of all three people
local_performance_df = global_df.pivot(index='Date', columns='level_0', values='Grade')
local_performance_df.index.name = 'Date'  # Set an appropriate name for the index
local_performance_df.columns.name = 'Person'  # Set an appropriate name for the columns
ax = local_performance_df.plot(kind='bar', figsize=(18, 6), alpha=1, color=bar_colors)
plt.title('Local Performance')
plt.xlabel('')
plt.ylabel('Mark')

ax.xaxis_date()  # Interpret the x-axis values as dates

# Replace x-axis numbers with exam names from the "Name" column
exam_dates = global_df['Date'].unique()
ax.set_xticks(range(len(exam_dates)))
ax.set_xticklabels([date.strftime('%b %Y') for date in exam_dates], rotation=45, fontsize=8)

# Add Cumulative Weighted Mean lines for each person
for sheet_index, sheet_name in enumerate(sheet_names):
    cumulative_means = cumulative_weighted_means_data[sheet_index]['Cumulative Weighted Mean']
    x_values = list(range(len(cumulative_means)))
    plt.plot(x_values, cumulative_means, linestyle='-', label=f'{sheet_name}', color=line_colors[sheet_index])

# Set the grid to be below the bars
ax.set_axisbelow(True)
ax.grid(color='gray')

plt.ylim(17, 31)

# Create a bar legend for the bars
handles, labels = ax.get_legend_handles_labels()
bar_legend = plt.legend(handles[:len(sheet_names)], labels[:len(sheet_names)], title='Weighted Avg Mark', loc='upper right')

# Create a line legend for the lines
line_legend = plt.legend(handles[len(sheet_names):], labels[len(sheet_names):], title='Mark', loc='upper left')

# Adjust the location of the line legend
plt.gca().add_artist(line_legend)

# Add the legends to the plot
plt.gca().add_artist(bar_legend)

# Specify the file path where you want to save the plot
output_file = os.path.join(current_directory,'../img/local_performance_plot.png')

# Save the plot to the specified file
fig = plt.gcf()
fig.set_size_inches((20 , 11.25), forward=False) #1920x1080
fig.savefig(output_file, dpi=500, bbox_inches='tight')
plt.show()

# Display the saved file path
print(f"Plot saved to '{output_file}'")
