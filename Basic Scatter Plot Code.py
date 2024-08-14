import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.signal import butter
from scipy.signal import filtfilt
from matplotlib.widgets import CheckButtons
from scipy.signal import find_peaks

####### Read the Excel file #######
excel_file = pd.ExcelFile('/Users/akin.o.akinjogbin/Desktop/July 2023 Tine Testing Raw data/Tine Engagement /Proccessed Files/Overview/ProcessedData_v7 vs v8.xlsx')

# Create a figure and a single subplot
fig, ax = plt.subplots(figsize=(20, 20))

# Create an empty DataFrame to store the peak coordinates
peak_data = pd.DataFrame(columns=['Sheet', 'Initial Peak X (Filtered)', 'Initial Peak Y (Filtered)', 'Max Peak X (Filtered)', 'Max Peak Y (Filtered)'])

# Create a list to store the visibility status of each line
line_visibility = []

# Iterate over each sheet in the workbook
for i, sheet_name in enumerate(excel_file.sheet_names):
    # Read the sheet data
    data = pd.read_excel(excel_file, sheet_name=sheet_name)

    # Specify column numbers for x and y
    x_column = 2  # Column number for x (starting from 0)
    y_column = 0  # Column number for y (starting from 0)

    # Extract the x and y columns from the sheet data
    x = data.iloc[:, x_column].values
    y = data.iloc[:, y_column].values

    # Convert y values to absolute values
    y = np.abs(y)

    # Set the filter order and cutoff frequency
    filter_order = 5  # Adjust the filter order as needed
    cutoff_freq = 0.01  # Adjust the cutoff frequency as needed

    # Apply Butterworth filtering to the data
    b, a = butter(filter_order, cutoff_freq, fs=1, btype='low', analog=False)
    filtered_y = filtfilt(b, a, y)

    # Store the filtered data in the DataFrame
    data['Filtered Y'] = filtered_y

    # Plot the filtered data and store the line object
    line, = ax.plot(x, filtered_y, label=f'Paddle ({sheet_name})')
    line_visibility.append(line.get_visible())


font1 = {'family':'monospace','color':'Green','size':20,'fontweight':'bold'}
font2 = {'family':'monospace','color':'Green','size':12,'fontweight':'bold'}
font3 = {'family':'monospace','size':10,'fontweight':'bold'}

# Set plot labels and title
ax.set_xlabel('Distance (mm)',fontdict=font2)
ax.set_ylabel('Force(N)',fontdict=font2)
ax.set_title('Paddle Engagement Comparison -  Version 7 vs Version 8',fontdict=font1)

# Function to handle legend checkbox updates
def update_legend(label):
    index = legend_labels.index(label)
    line_visibility[index] = not line_visibility[index]
    ax.get_lines()[index].set_visible(line_visibility[index])
    plt.draw()

# Create a list to store the legend labels
legend_labels = [line.get_label() for line in ax.get_lines()]

######## Create the CheckButtons instance and connect it to the update_legend function. Turn off if needed.###########
#checkbox_ax = fig.add_axes([0.1, 0.6, 0.3, 0.3])  # Adjust the position and size as needed
#check_buttons = CheckButtons(checkbox_ax, labels=legend_labels, actives=line_visibility)
#check_buttons.on_clicked(update_legend)

# Add the legend to the figure
ax.legend()


# Show the plot
plt.show()
