import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

font1 = {'family': 'monospace', 'color': 'Green', 'size': 20, 'fontweight': 'bold'}
font2 = {'family': 'monospace', 'color': 'Green', 'size': 12, 'fontweight': 'bold'}
font4 = {'family':'monospace','color':'black','size':10,'fontweight':'bold'}
font3 = {'family': 'monospace', 'size': 10, 'fontweight': 'bold'}
font5 = {'family':'monospace','color':'red','size':12,'fontweight':'bold'}

# Add a title and labels to the axes
plt.title('Cuff Position', fontdict=font1)
plt.xlabel('Galvani Lots', fontdict=font2)
plt.ylabel('Force (Cm)', fontdict=font2)

# Read data from the Excel file (replace 'data.xlsx' with your file name)
data_df = pd.read_excel('/Users/akin.o.akinjogbin/Library/CloudStorage/GoogleDrive-akin.o.akinjogbin@galvani.bio/Shared drives/Neural Interfaces/Splenic Lead/Complaints and CAPAs/CAPA-0035/NCM Report Report 55869.xlsx', sheet_name='Cufff Postional Data')  

# Dictionary to map column names to custom colors
# CURRENTLY SET UP FOR 8 COLUMNS
column_colors = {
    'Column1': 'blue',
    'Column2': 'green',
    'Column3': 'red',
    'Column4': 'purple',
    'Column5': 'brown',
    'Column6': 'grey',
    'Column7': 'orange',
    'Column8': 'magenta',
    # Add more columns and their colors as needed
}

# Create the box plot using Seaborn and assign custom colors
ax = sns.boxplot(data=data_df, width=0.35, palette=column_colors.values(), whis=1.5)  # Adjust whis as needed

# Calculate mean values for each column
mean_values = data_df.mean()
max_values = data_df.max()

# Add mean values as points and annotations above the max line
for i, mean_val in enumerate(mean_values):
    ax.plot(i, mean_values[i], marker='x', markersize=8, color='black', linestyle='None')
    ax.text(i, max_values[i] + 0.01, f'Mean: {mean_val:.2f}', ha='center', fontdict=font5)

# Set the legend for the colors
#legend_list = [plt.Rectangle((0, 0), 1, 1, fc=color) for color in column_colors.values()]
#plt.legend(legend_list, data_df.columns, loc='lower right')

# Show the plot
plt.show()
