from heapq import nsmallest
import logging as log
from sqlite3 import Timestamp
import pandas as pd
import numpy as np
import os 
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import glob

dir_path = os.path.dirname(os.path.realpath(__file__))

def process_file(filename):
    try:
        file = pd.read_excel(filename).dropna(subset=['active_power_im'])
        line_chart = filename.replace("files","results").replace(".xlsx",".png")
        # delete data by condition
        file = file[file['active_power_im'] != "data_faltante"]

        inversores = file['id_i'].unique() 

        fig, ax = plt.subplots(figsize=(16, 8))

        text_results = []

        for i, inversor in enumerate(inversores):
            ax.plot('fecha_im', 'active_power_im', data=file.loc[file.id_i==inversor, :], label=inversor)
            text_results.append({
                'inversor': inversor,
                'total': int(file.loc[file['id_i'] == inversor, 'active_power_im'].sum()),
                'ruta': line_chart, 
                'min': 0,
                'max': 0
            })

            #print(np.where((file['id_i'] == inversor), file['active_power_im']).max(1))


        #print(text_results)

        fecha_excel = file['fecha_im'].tolist()[0]

        hours = mdates.HourLocator(interval = 1)
        h_fmt = mdates.DateFormatter('%H:%M')

        ax.xaxis.set_major_locator(hours)
        ax.xaxis.set_major_formatter(h_fmt)

        plt.ylabel('$Active$ $Power$')
        plt.yticks(fontsize=12, alpha=.7)
        plt.title("Energía Solar Generada el " + fecha_excel.strftime('%d de %B del %Y'), fontsize=22)
        plt.grid(axis='y', alpha=.3)

        # Remove borders
        plt.gca().spines["top"].set_alpha(0.0)    
        plt.gca().spines["bottom"].set_alpha(0.0)
        plt.gca().spines["right"].set_alpha(0.0)    
        plt.gca().spines["left"].set_alpha(0.0)   
        plt.legend(loc='upper right', ncol=2, fontsize=12)

        plt.savefig(line_chart)

        with open(line_chart.replace('png', 'txt'), 'w') as f:
            for results in text_results:
                f.write(
                    "Inversionista: {inversor}\nTotal Power Active: {total}\nMin Power Energy: {min}\nMax Power Energy: {max}\nRuta de Gráfica: {ruta}\n\n".format(**results)
                )


        print("Total Active Power de todas las plantas: ", str(sum([int(i) for i in file['active_power_im'].tolist() if isinstance(i, float) or isinstance(i, int)])))
        
    except Exception as e:
        log.error(e)

if __name__=="__main__":
    files = glob.glob(r'{0}/files/*.xlsx'.format(dir_path))
    for file in files:
        process_file(file)