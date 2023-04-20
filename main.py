# Importing needed libraries
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
import plotly.graph_objects as go

# Read file and sheets
dataset = pd.ExcelFile('Dataset.xls')
tbl1_HHtype = pd.read_excel(dataset, 'Tbl 1 HH Type').reset_index(drop=True)
tbl2_tbl15_UseCar_HaveCar = pd.read_excel(dataset, 'Tbl 2 Use Car & Tbl 15 Have car')
tbl4_AreaType = pd.read_excel(dataset, 'Tbl 4 Area Type')
tbl6_Area = pd.read_excel(dataset, 'Tbl 6 Area')
tbl10_AgeCohort = pd.read_excel(dataset, 'Tbl 10 Age cohort')
tbl11_Employment = pd.read_excel(dataset, 'Tbl 11 Employment')
tbl35_HSP_HHType = pd.read_excel(dataset, 'Tbl 35 Gttng frm-to HSP HH type')
tbl37_HSP_Employment = pd.read_excel(dataset, 'Tbl 37 Gttng from-to HSP Emplmt')
tbl39_tbl41_SafeAfterDark = pd.read_excel(dataset, 'Tbl 39+41 Fling safe after drk')

# Tbl 1

fig, axs = plt.subplots(1, sharex=True, sharey=True)  # Initiating
plt.xticks(fontsize=4)  # Changing font size to fit labels
fig.suptitle('Mean satisfaction by Household Type')  # Title

g1 = axs.bar(tbl1_HHtype['Category'], tbl1_HHtype['Mean'])  # Bar chart
axs.bar_label(g1)  # Showing data points
plt.show()  # Showing chart

# Tbl 2 & Tbl 15

# ------ Creating the pie chart labels
temp = tbl2_tbl15_UseCar_HaveCar.iloc[:, 4].to_list()
temp2 = tbl2_tbl15_UseCar_HaveCar.iloc[:, 0].to_list()
lbl = list()
for i, j in zip(temp, temp2):
    lbl.append(j + '\n' + str(i * 100) + '%')

# ----- Making chart
fig, axs = plt.subplots(2)  # Initiating
axs[0].pie(tbl2_tbl15_UseCar_HaveCar['% have use of car'], labels=lbl, colors=['blue', 'orange'])  # Pie chart
axs[0].set_title('Having use of car')  # Pie chart title
g1 = axs[1].bar(tbl2_tbl15_UseCar_HaveCar['Use of car'], tbl2_tbl15_UseCar_HaveCar['Satisfaction Mean'],
                color=['blue', 'orange'])  # Bar chart
axs[1].set_title('Satisfaction mean by having use of car')  # Bar chart title
axs[1].bar_label(g1)  # Bar chart data points

plt.show()

# Tbl 4 & Tbl 6

temp = tbl4_AreaType[(tbl4_AreaType.Category == "Urban")]['Mean']  # Creating temp table
temp = temp.to_list()  # Changing df to list
temp = [round(num, 1) for num in temp]  # Rounding

fig = go.Figure(go.Indicator(
    mode="gauge+number",
    value=temp[0],
    domain={'x': [0, 1], 'y': [0, 1]},
    title={'text': "Mean satisfaction in Urban areas"},
    gauge={'axis': {'range': [0, 10]}}))  # Creating gauge

fig.show()

temp = tbl4_AreaType[(tbl4_AreaType.Category == "Rural")]['Mean']  # Creating temp table
temp = temp.to_list()  # Changing df to list
temp = [round(num, 1) for num in temp]  # Rounding

fig = go.Figure(go.Indicator(
    mode="gauge+number",
    value=temp[0],
    domain={'x': [0, 1], 'y': [0, 1]},
    title={'text': "Mean satisfaction in Rural areas"},
    gauge={'axis': {'range': [0, 10]}}))  # Creating gauge

fig.show()

# Tbl 10

fig, axs = plt.subplots(1, sharex=True, sharey=True)  # Initiating
plt.xticks(fontsize=10)  # Changing font size to fit labels
fig.suptitle('Satisfaction by age cohort')  # Title
g1 = axs.bar(tbl10_AgeCohort['Age cohort'], tbl10_AgeCohort['Mean'])  # Bar chart
axs.bar_label(g1)  # Bar chart data points

plt.show()  # Showing chart

# Tbl 11

temp = tbl11_Employment[(tbl11_Employment['Employment Status'] == "In employment")]['Mean']  # Creating temp table
temp = temp.to_list()  # Changing df to list
temp = [round(num, 1) for num in temp]  # Rounding

fig = go.Figure(go.Indicator(
    mode="gauge+number",
    value=temp[0],
    domain={'x': [0, 1], 'y': [0, 1]},
    title={'text': "Mean satisfaction for employed"},
    gauge={'axis': {'range': [0, 10]}}))  # Creating gauge

fig.show()

temp = tbl11_Employment[(tbl11_Employment['Employment Status'] == "Not in employment")]['Mean']  # Creating temp table
temp = temp.to_list()  # Changing df to list
temp = [round(num, 1) for num in temp]  # Rounding

fig = go.Figure(go.Indicator(
    mode="gauge+number",
    value=temp[0],
    domain={'x': [0, 1], 'y': [0, 1]},
    title={'text': "Mean satisfaction for not employed"},
    gauge={'axis': {'range': [0, 10]}}))  # Creating gauge

fig.show()

# Tbl 35

# Creating needed DFs from the main df
very_easy = tbl35_HSP_HHType[(tbl35_HSP_HHType.Classification == "Very easy")]  # Very easy tbl
very_easy = very_easy.round(1)  # Rounding
fairly_easy = tbl35_HSP_HHType[(tbl35_HSP_HHType.Classification == "Fairly easy")]  # Fairly easy tbl
fairly_easy = fairly_easy.round(1)  # Rounding
fairly_difficult = tbl35_HSP_HHType[(tbl35_HSP_HHType.Classification == "Fairly difficult")]  # Fairly difficult tbl
fairly_difficult = fairly_difficult.round(1)  # Rounding
very_difficult = tbl35_HSP_HHType[(tbl35_HSP_HHType.Classification == "Very difficult")]  # Very difficult tbl
very_difficult = very_difficult.round(1)  # Rounding
place = np.arange(7)  # Numbering for labels
categories = very_easy['Category'].to_list()  # Available categories

fig, ax = plt.subplots()  # Initiating
g1 = ax.bar(place - 0.4, very_easy["%"], 0.2, color="blue", label="Very easy")  # Firt bar group (very easy)
g2 = ax.bar(place - 0.2, fairly_easy['%'], 0.2, color="cyan", label="Fairly easy")  # Second bar group (fairly easy)
g3 = ax.bar(place, fairly_difficult['%'], 0.2, color='orange',
            label="Fairly difficult")  # Third bar group (fairly difficult)
g4 = ax.bar(place + 0.2, very_difficult['%'], 0.2, color="red",
            label="Very difficult")  # Fourth bar group (very difficult)

ax.set_xticks(place, categories, fontsize=4)  # Setting the categories
ax.set_ylabel("% of responses")  # Setting the title of y axis
ax.legend(["Very easy", "Fairly easy", "Fairly difficult", "Very difficult"])  # Setting the legend
ax.set_title("Ease of getting to and from the hospital by household type")  # Setting the chart title
ax.bar_label(g1, padding=3)  # First bar group data points
ax.bar_label(g2, padding=3)  # Second bar group data points
ax.bar_label(g3, padding=3)  # Third bar group data points
ax.bar_label(g4, padding=3)  # Fourth bar group data points

plt.show()

# Tbl 37

# Creating needed DFs from the main df
inemployment = tbl37_HSP_Employment[(tbl37_HSP_Employment.Category == "In employment")]  # In employment tbl
inemployment = inemployment.round(1)  # Rounding
outemployment = tbl37_HSP_Employment[(tbl37_HSP_Employment.Category == "Not in employment")]  # Not in employment tbl
outemployment = outemployment.round(1)  # Rounding

place = np.arange(4)  # Numbering for labels
categories = inemployment['Classification'].to_list()  # Available classifications
fig, ax = plt.subplots()  # Initiating
g1 = ax.bar(place - 0, inemployment['%'], 0.2, color="blue")  # In employment bar group
g2 = ax.bar(place + 0.2, outemployment['%'], 0.2, color="cyan")  # Not in employment bar group

ax.set_xticks(place, categories, fontsize=10)  # Setting x axis categories
ax.set_ylabel("% of responses")  # Setting yaxis title
ax.legend(["In employment", "Out of employment"])  # Setting legend
ax.set_title("Ease of getting to and from hospital, by employment status")  # Setting chart title
ax.bar_label(g1, padding=3)  # Setting in employment chart data points
ax.bar_label(g2, padding=3)  # Setting not in employment chart data points

plt.show()

# Tbl 39 & Tbl 41

# Creating needed DFs from the main df
male = tbl39_tbl41_SafeAfterDark[['Men %', 'Category']].copy()  # Male tbl
female = tbl39_tbl41_SafeAfterDark[['Women %', 'Category']].copy()  # Female tbl

fig, ax = plt.subplots(ncols=2, sharey=True)  # initiating

g1 = ax[0].barh(range(0, len(tbl39_tbl41_SafeAfterDark)), male['Men %'], align="center", color="blue")  # Men bar group
ax[0].set(title="Male")  # Setting men bar group title
g2 = ax[1].barh(range(0, len(tbl39_tbl41_SafeAfterDark)), female['Women %'], align="center",
                color="pink")  # Women bar group
ax[1].set(title="Female")  # Setting women bar group title
fig.suptitle("% Feeling of safety travelling by public transport after dark")  # Setting chart title
ax[0].set(yticks=range(0, len(tbl39_tbl41_SafeAfterDark)), yticklabels=male['Category'])  # Setting y axis labels
ax[0].bar_label(g1, padding=3)  # Setting men bar group data points
ax[1].bar_label(g2, padding=3)  # Setting women bar group data points
ax[0].invert_xaxis()  # Inverting axis to create population pyramid effect

plt.show()
