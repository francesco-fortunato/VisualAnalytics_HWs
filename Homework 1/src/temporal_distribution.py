import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import matplotlib.dates as mdates  # Import the mdates module for date formatting
import os  

# Define a custom color palette
custom_palette = sns.color_palette("tab20")  # You can choose other palettes

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
global_std_data = []

# Define a list of colors for Cumulative Weighted Mean lines
bar_colors = custom_palette[:len(sheet_names)]
line_colors = custom_palette[len(sheet_names):]

# Create a list to store Cumulative Weighted Mean data for each person
cumulative_weighted_means_data = []

# Iterate through each person's data
for sheet_index, sheet_name in enumerate(sheet_names):
    if (sheet_name=='Francesco'):
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
        cumulative_std = []

        # Iterate through the rows and calculate cumulative weighted mean and standard deviation
        for index, row in df.iterrows():
                cumulative_weighted_sum += row['Grade'] * row['Credits']
                cumulative_credits_sum += row['Credits']
                cumulative_weighted_mean = cumulative_weighted_sum / cumulative_credits_sum
                cumulative_weighted_means.append(cumulative_weighted_mean)
                cumulative_std.append(df['Grade'].std())
        
        # Add the cumulative weighted mean and standard deviation to the DataFrame
        df['Cumulative Weighted Mean'] = cumulative_weighted_means
        df['Cumulative Std Deviation'] = cumulative_std

        # Plot the temporal distribution with cumulative weighted mean and standard deviation
        plt.figure(figsize=(10, 6))
        plt.plot(df['Date'], df['Grade'], marker='o', markersize=3, linestyle='-', color=custom_palette[0])
        # plt.axhline(cumulative_weighted_means[-1], label=f'Weighted Grade Mean', linestyle='--', color=line_colors[sheet_index])
        plt.title(f'Francesco Fortunato - Temporal Distribution of Exam Scores')
        plt.xlabel('')
        plt.ylabel('Grade')
        plt.grid(True)
        ax = plt.gca()

        ax.fill_between([], [], color=custom_palette[1], alpha=0.5, label='Exam Sessions')

        # Create shaded regions for January and February for each year
        for year in range(2019, 2023):
                jan_feb_start = pd.Timestamp(f'{year}-01-01')
                jan_feb_end = pd.Timestamp(f'{year}-02-28')
                ax.fill_between([jan_feb_start, jan_feb_end], 15, 35, alpha=0.5, color=custom_palette[1])
        for year in range(2019, 2023):
                jan_feb_start = pd.Timestamp(f'{year}-06-01')
                jan_feb_end = pd.Timestamp(f'{year}-07-31')
                ax.fill_between([jan_feb_start, jan_feb_end], 15, 35, alpha=0.5, color=custom_palette[1])
        for year in range(2019, 2022):
                jan_feb_start = pd.Timestamp(f'{year}-09-01')
                jan_feb_end = pd.Timestamp(f'{year}-09-30')
                ax.fill_between([jan_feb_start, jan_feb_end], 15, 35, alpha=0.5, color=custom_palette[1])
        
        plt.ylim(17, 31)

        x_ticks = pd.date_range(start='2019-01-01', end='2022-08-01', freq='2MS')
        ax.set_xticks(x_ticks)
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%b %Y'))  # Format x-tick labels

        plt.xticks(rotation=45)
        
        # Print and display local performance
        print(f'{sheet_name} - Cumulative Weighted Mean (Last Entry): {cumulative_weighted_means[-1]:.2f}')
        print(f'{sheet_name} - Cumulative Std Deviation (Last Entry): {cumulative_std[-1]:.2f}')
        
        dataframes.append(df)

        plt.legend()

        
        # Calculate and store global performance data for this person
        average_grade = df['Grade'].mean()
        std_deviation = df['Grade'].std()
        global_performance_data.append({
                'Person': sheet_name,
                'Average Grade': average_grade
        })
        global_std_data.append({
                'Person': sheet_name,
                'Std Deviation': std_deviation
        })

        # Store Cumulative Weighted Mean data for each person
        cumulative_weighted_means_data.append({
                'Person': sheet_name,
                'Cumulative Weighted Mean': cumulative_weighted_means
        })

        if(sheet_index==1):
                # Here: add a new tick with the required value
                ax = plt.gca()

                yticks = [18, 20, 22, 24, 26, 28, 30]
                yticklabels = [*ax.get_yticklabels()]
                ax.set_yticks(yticks)
                ax.set_axisbelow(True)

                
                fig = plt.gcf()
                fig.set_size_inches((20 , 11.25), forward=False) #1920x1080
                fig.savefig(os.path.join(current_directory, '../img/temporal_distribution.png'), dpi=500, bbox_inches='tight')

plt.show()