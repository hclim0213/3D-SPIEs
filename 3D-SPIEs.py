class PlotlyVisualizer:
    @staticmethod
    def excel_reader(fileName, Sheet_name, listOfDroppingIndices):
        from pandas import read_excel
        from pandas import read_csv

        df_Sheet = read_excel(fileName, index_col=0, header=0, sheetname=Sheet_name)

        df_Sheet.index = df_Sheet.index.str.encode('utf-8')
        df_Sheet.columns = df_Sheet.columns.str.encode('utf-8')

        if len(listOfDroppingIndices) != 0:
            df_Sheet = df_Sheet.drop(listOfDroppingIndices, axis=0)

        return df_Sheet


def visualizer1(x_axisTitle, y_axisTitle, z_axisTitle, df_Single_Energy='', df_Single_Distance='', ES_sheet='',
                EX_sheet='', CT_sheet='', DI_sheet='', SL_sheet='', energySheet='', distanceSheet='', boolean_PIEDA="0",
                boolean_Sol="0", COMMON_NAME=''):
    import plotly.offline as py
    import plotly.graph_objs as go

    def color_switcher(color_table, energy):
        energy_category = [-20, -15, -7, -5, -3, 0, 3, 5, 7, 15, 20]
        rdbu_reverse = ['rgb(5,48,97)', 'rgb(33,102,172)', 'rgb(67,147,195)', 'rgb(146,197,222)',
                        'rgb(209,229,240)', 'rgb(247,247,247)', 'rgb(253,219,199)', 'rgb(244,165,130)',
                        'rgb(214,96,77)', 'rgb(178,24,43)', 'rgb(103,0,31)']
        rdbu2_reverse = ['rgb(5,48,97)', 'rgb(33,102,172)', 'rgb(67,147,195)', 'rgb(146,197,222)',
                         'rgb(209,229,240)', '#FFFFFF', 'rgb(253,219,199)', 'rgb(244,165,130)',
                         'rgb(214,96,77)', 'rgb(178,24,43)', 'rgb(103,0,31)']
        rdylbu_reverse = ['rgb(49,54,149)', 'rgb(69,117,180)', 'rgb(116,173,209)', 'rgb(171,217,233)',
                          'rgb(224,243,248)', 'rgb(255,255,191)', 'rgb(254,224,144)', 'rgb(253,174,97)',
                          'rgb(244,109,67)', 'rgb(215,48,39)', 'rgb(165,0,38)']
        if color_table is 'rdbu':
            color_table_list = rdbu_reverse
        elif color_table is 'rdbu2':
            color_table_list = rdbu2_reverse
        elif color_table is 'rdylbu':
            color_table_list = rdylbu_reverse

        if energy < energy_category[0]:
            return color_table_list[0]
        elif energy_category[0] <= energy and energy < energy_category[1]:
            return color_table_list[1]
        elif energy_category[1] <= energy and energy < energy_category[2]:
            return color_table_list[2]
        elif energy_category[2] <= energy and energy < energy_category[3]:
            return color_table_list[3]
        elif energy_category[3] <= energy and energy < energy_category[4]:
            return color_table_list[4]
        elif energy_category[4] <= energy and energy < energy_category[5]:
            return color_table_list[5]
        elif energy_category[5] <= energy and energy < energy_category[6]:
            return color_table_list[5]
        elif energy_category[6] <= energy and energy < energy_category[7]:
            return color_table_list[6]
        elif energy_category[7] <= energy and energy < energy_category[8]:
            return color_table_list[7]
        elif energy_category[8] <= energy and energy < energy_category[9]:
            return color_table_list[8]
        elif energy_category[9] <= energy and energy < energy_category[10]:
            return color_table_list[9]
        else:
            return color_table_list[10]

    color_table_name = 'rdbu'
    list_Figures = []
    listOfFragment1 = []
    listOfFragment2 = []
    listOfEnergy_color = []
    listOfDistance = []
    listOfEnergy = []
    value = []

    for fragment1 in df_Single_Distance.index.tolist():
        for fragment2 in df_Single_Distance.columns.tolist():
            listOfFragment1.append(fragment1)
            listOfFragment2.append(fragment2)
            listOfEnergy_color.append(color_switcher(color_table_name, df_Single_Energy.loc[fragment1, fragment2]))
            listOfEnergy.append(df_Single_Energy.loc[fragment1, fragment2])
            listOfDistance.append(df_Single_Distance.loc[fragment1, fragment2])
            INFO = "x: " + fragment1
            INFO = INFO + '<br>' + "y: " + fragment2
            INFO = INFO + '<br>' + "Energy: " + str(df_Single_Energy.loc[fragment1, fragment2]) + " kcal/mol"
            INFO = INFO + '<br>' + "Distance: {:.3f} Ang.".format(df_Single_Distance.loc[fragment1, fragment2])
            if "1" in boolean_PIEDA:
                INFO = INFO + '<br>' + "ES: " + str(ES_sheet.loc[fragment1, fragment2]) + " kcal/mol"
                INFO = INFO + '<br>' + "EX: " + str(EX_sheet.loc[fragment1, fragment2]) + " kcal/mol"
                INFO = INFO + '<br>' + "CT: " + str(CT_sheet.loc[fragment1, fragment2]) + " kcal/mol"
                INFO = INFO + '<br>' + "DI: " + str(DI_sheet.loc[fragment1, fragment2]) + " kcal/mol"
            if "1" in boolean_Sol:
                INFO = INFO + '<br>' + "SL: " + str(SL_sheet.loc[fragment1, fragment2]) + " kcal/mol"

            if "1" in boolean_PIEDA:
                if (float(ES_sheet.loc[fragment1, fragment2]) != 0):
                    if "1" in boolean_Sol:
                        INFO = INFO + '<br>' + "EaS: " + "1 : " + str(
                            round(float(EX_sheet.loc[fragment1, fragment2]) / float(ES_sheet.loc[fragment1, fragment2]),
                                  3))
                        INFO = INFO + " : " + str(
                            round(float(CT_sheet.loc[fragment1, fragment2]) / float(ES_sheet.loc[fragment1, fragment2]),
                                  3))
                        INFO = INFO + " : " + str(
                            round(float(DI_sheet.loc[fragment1, fragment2]) / float(ES_sheet.loc[fragment1, fragment2]),
                                  3))
                        INFO = INFO + " : " + str(
                            round(float(SL_sheet.loc[fragment1, fragment2]) / float(ES_sheet.loc[fragment1, fragment2]),
                                  3))
                        INFO = INFO + '<br>' + "= ES : EX : CT : DI : SL"
                    else:
                        INFO = INFO + '<br>' + "EaS: " + "1 :" + str(
                            round(float(EX_sheet.loc[fragment1, fragment2]) / float(ES_sheet.loc[fragment1, fragment2]),
                                  3))
                        INFO = INFO + " : " + str(
                            round(float(CT_sheet.loc[fragment1, fragment2]) / float(ES_sheet.loc[fragment1, fragment2]),
                                  3))
                        INFO = INFO + " : " + str(
                            round(float(DI_sheet.loc[fragment1, fragment2]) / float(ES_sheet.loc[fragment1, fragment2]),
                                  3))
                        INFO = INFO + '<br>' + "= ES : EX : CT : DI"

            value.append(INFO)

    x, y, z, color = listOfFragment1, listOfFragment2, listOfDistance, listOfEnergy_color

    list_Figures.append(go.Scatter3d(
        x=x,
        y=y,
        z=z,
        text=value,
        hoverinfo="text",
        mode='markers',
        marker=dict(
            color=color,
            cauto=False,
            cmax=25,
            cmin=-25,
            size=6,
            symbol='circle',
            line=dict(
                color=color,
                width=1
            ),
            showscale=True,
            #            colorscale=[[-20, 'rgb(5,48,97)'], [-15, 'rgb(33,102,172)'], [-7, 'rgb(67,147,195)'], [-5, 'rgb(146,197,222)'],
            #                        [-3, 'rgb(209,229,240)'], [0, 'rgb(247,247,247)'], [3, 'rgb(253,219,199)'], [5, 'rgb(244,165,130)'],
            #                        [7, 'rgb(214,96,77)'], [15, 'rgb(178,24,43)'], [20, 'rgb(103,0,31)']],
            colorbar=dict(
                title='ENERGY',
                lenmode='pixels',
                len=600,
                titleside='top',
                tickmode='array',
                tickvals=[-22.5, -20, -15, -7, -5, -3, 0, 3, 5, 7, 15, 20, 22.5],
                ticktext=['UNDER', '-20', '-15', '-7', '-5', '-3', '0', '3', '5', '7', '15', '20', 'OVER'],
                ticks='outside'
            )
        ),
        opacity=0 if color is 'rgb(247,247,247)' else 0.6,
    ))

    data = list_Figures
    layout = {
        "scene": {
            "xaxis": {
                "title": x_axisTitle
            },
            "yaxis": {
                "title": y_axisTitle
            },
            "zaxis": {
                "title": z_axisTitle
            }
        }
    }
    fig = go.Figure(data=data, layout=layout)

    if COMMON_NAME != '':
        py.plot(fig, filename='3D-SPIEs' + '_' + COMMON_NAME + '.html')
    else:
        py.plot(fig, filename='3D-SPIEs' + '_E_' + energySheet + '_D_' + distanceSheet + '.html')

print "------------------------------------------------------------------------------------------------------------"
print "The program makes the 3D Scattered Plot with Colored Dots in plotly from your Distance and Energy excel file"
print "Author : Jungho chun and Hocheol Lim"
print "Only permit to use this program and the other actions including distribution and copy are not permitted"
print "All rights are reserved by copyright"
print "------------------------------------------------------------------------------------------------------------"

# Initiation of variables
boolean_PIEDA = "0"
boolean_Solvation = "0"
boolean_Automode = "0"
boolean_detailedMode = "1"
COMMON_NAME = ''

dirFILE = raw_input("Type the directory of FILE: ")
boolean_PIEDA = raw_input("Select PIEDA-MODE. (1/0) :")
boolean_Solvation = raw_input("Select Solvation-MODE. (1/0) :")
boolean_Automode = raw_input("Select Auto-mode 1 for inserting each name of sheets. (1/0)")

print ""
if "1" in boolean_Automode:
    print "---------------------------------------------------------------"
    print "Auto requires fixed form for name of sheet."
    print "You should insert only COMMON NAME"
    print "Energy sheet : COMMON NAME_Total"
    print "Distance sheet : COMMON NAME_Distance"
    print "Electrostatic sheet : COMMON NAME_ES"
    print "Exchange-Repulsion sheet : COMMON NAME_EX"
    print "Charge Transfer sheet : COMMON NAME_CT"
    print "Dispersion sheet : COMMON NAME_DI"
    print "Solvation sheet : COMMON NAME_SL"
    print "---------------------------------------------------------------"
    COMMON_NAME = raw_input("Type COMMON NAME: ")
    Energy_sheet = COMMON_NAME + "_Total"
    Distance_sheet = COMMON_NAME + "_Distance"
    if "1" in boolean_PIEDA:
        ES_sheet = COMMON_NAME + "_ES"
        EX_sheet = COMMON_NAME + "_EX"
        CT_sheet = COMMON_NAME + "_CT"
        DI_sheet = COMMON_NAME + "_DI"
        if "1" in boolean_Solvation:
            SL_sheet = COMMON_NAME + "_SL"
    else:
        if "1" in boolean_Solvation:
            SL_sheet = COMMON_NAME + "_SL"
else:
    Energy_sheet = raw_input("Type name of Energy sheet you want to plot: ")
    Distance_sheet = raw_input("Type name of Distance sheet you want to plot: ")
    if "1" in boolean_PIEDA:
        print "You selected PIEDA-MODE."
        ES_sheet = raw_input("Type name of Electrostatic sheet you want to plot: ")
        EX_sheet = raw_input("Type name of Exchange-Repulsion sheet you want to plot: ")
        CT_sheet = raw_input("Type name of Charge Transfer sheet you want to plot: ")
        DI_sheet = raw_input("Type name of Dispersion sheet you want to plot: ")
        if "1" in boolean_Solvation:
            print "You also selected Solvation-MODE, too."
            SL_sheet = raw_input("Type name of Solvation sheet you want to plot: ")
    else:
        if "1" in boolean_Solvation:
            print "You selected Solvation-MODE."
            SL_sheet = raw_input("Type name of Solvation sheet you want to plot: ")

print ""
print "---------------------------------------------------------------"
print "Titles of axes have next three options"
print "2 : Insert names of x, y, z axes."
print "1 : Default x, y, z axes"
print "0 : Nothing"
print "---------------------------------------------------------------"
boolean_detailedMode = raw_input("What option will you select for titles of axes? (2/1/0): ")
x_axisTitle = 'x'
y_axisTitle = 'y'
z_axisTitle = 'z'

if "2" in boolean_detailedMode:
    x_axisTitle = raw_input("Type name of x-axis: ")
    y_axisTitle = raw_input("Type name of y-axis: ")
    z_axisTitle = raw_input("Type name of z-axis: ")
if "0" in boolean_detailedMode:
    x_axisTitle = ''
    y_axisTitle = ''
    z_axisTitle = ''

visualized = PlotlyVisualizer()
print ""
print "---------------------------------------------------------------"
print "Bringing data from the input file..."

df_Single_Energy = visualized.excel_reader(fileName=dirFILE, Sheet_name=Energy_sheet, listOfDroppingIndices=[])
df_Single_Distance = visualized.excel_reader(fileName=dirFILE, Sheet_name=Distance_sheet, listOfDroppingIndices=[])

# Initiate list of PIEDA
df_Single_ES = ''
df_Single_EX = ''
df_Single_CT = ''
df_Single_DI = ''
df_Single_SL = ''

if "1" in boolean_PIEDA:
    df_Single_ES = visualized.excel_reader(fileName=dirFILE, Sheet_name=ES_sheet, listOfDroppingIndices=[])
    df_Single_EX = visualized.excel_reader(fileName=dirFILE, Sheet_name=EX_sheet, listOfDroppingIndices=[])
    df_Single_CT = visualized.excel_reader(fileName=dirFILE, Sheet_name=CT_sheet, listOfDroppingIndices=[])
    df_Single_DI = visualized.excel_reader(fileName=dirFILE, Sheet_name=DI_sheet, listOfDroppingIndices=[])

if "1" in boolean_Solvation:
    df_Single_SL = visualized.excel_reader(fileName=dirFILE, Sheet_name=SL_sheet, listOfDroppingIndices=[])

print("Done !!!")
print "---------------------------------------------------------------"

print ""
print "Making 3D-SPIEs from data..."
visualizer1(df_Single_Energy=df_Single_Energy, df_Single_Distance=df_Single_Distance,
            ES_sheet=df_Single_ES, EX_sheet=df_Single_EX, CT_sheet=df_Single_CT, DI_sheet=df_Single_DI,
            SL_sheet=df_Single_SL,
            energySheet=Energy_sheet, distanceSheet=Distance_sheet,
            boolean_PIEDA=boolean_PIEDA, boolean_Sol=boolean_Solvation,
            x_axisTitle=x_axisTitle, y_axisTitle=y_axisTitle, z_axisTitle=z_axisTitle, COMMON_NAME=COMMON_NAME)
print "Done !!!"
print "---------------------------------------------------------------"
