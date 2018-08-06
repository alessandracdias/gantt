import datetime as dt
import matplotlib.pyplot as plt
import matplotlib.font_manager as font_manager
import matplotlib.dates
from matplotlib.dates import WEEKLY, MONTHLY, DateFormatter, rrulewrapper, RRuleLocator
import numpy as np
from collections import namedtuple
import pandas as pd
from datetime import datetime


def CreateGanttChart(my_dic,  first, last):
    """
        Create gantt charts with matplotlib
        Give file name.
    """

    pos = np.arange(0.5, (len(my_dic) * 0.5) + 0.5, 0.5)
    fig = plt.figure(figsize=(100, 30))
    ax = fig.add_subplot(111)
    cont = 0

    for item in my_dic.items():
        cont = cont + 1
        n = len(item[1]['end_date'])
        for counter in range(0, n):

            curr_barh = ax.barh(((cont - 1) * 0.5) + 0.5, item[1]['width'][counter], left=item[1]['pos'][counter],
                                height=0.2, align='center', color='#80aaff', alpha=0.8)[0]

            text = item[1]['label'][counter]
            # Center the text vertically in the bar
            yloc = curr_barh.get_y() + (curr_barh.get_height() / 2.0)

            ax.text(item[1]['pos'][counter] + (item[1]['width'][counter]) / 2.0, yloc, text, fontsize=9, horizontalalignment='center',
                    verticalalignment='center', color='white', weight='bold',
                    clip_on=True)

    locsy, labelsy = plt.yticks(pos, my_dic.keys())
    plt.setp(labelsy, fontsize=10)
    ax.set_xlabel('Hours')
    ax.set_xticks(np.arange(np.floor(first), np.ceil(last), 1))
    ax.set_ylim(ymin=-0.05, ymax=len(my_dic.keys()) * 0.5 + 0.5)
    ax.grid(color='g', linestyle=':')
    rule = rrulewrapper(WEEKLY, interval=1)
    loc = RRuleLocator(rule)

    # formatter = DateFormatter("%d-%b")

    font = font_manager.FontProperties(size='small')
    ax.legend(loc=1, prop=font)

    ax.invert_yaxis()
    fig.autofmt_xdate()

    fig.tight_layout(rect=[0.027, 0.069, 0.984, 0.981], h_pad=0.2, w_pad=0.2)
    plt.savefig('gantt.svg')
    plt.show()


if __name__ == '__main__':

    xls = pd.ExcelFile("FlightLegsRepublicNovo.xls")

    sheetX = xls.parse(0)  # 2 is the sheet number

    acfts = sheetX['Acft'].values
    orig = sheetX['LegOrig'].values
    dest = sheetX['LegDest'].values
    leg_etd = sheetX['LegETD'].values
    leg_eta = sheetX['LegETA'].values

    my_dic = {}

    for index, actf in enumerate(acfts):
        if not actf in my_dic:
            my_dic[actf] = {'start_date': [], 'end_date': [],
                            'label': [], 'width': [], 'pos': []}

        ts = (leg_etd[index] - np.datetime64('1970-01-01T00:00:00Z')
              ) / np.timedelta64(1, 's')
        start_data = datetime.utcfromtimestamp(ts)

        ts = (leg_eta[index] - np.datetime64('1970-01-01T00:00:00Z')
              ) / np.timedelta64(1, 's')
        end_data = datetime.utcfromtimestamp(ts)

        duration = ((end_data - start_data).total_seconds() / 60.0) / 60
        my_dic[actf]['width'].append(duration)
        my_dic[actf]['start_date'].append(start_data)
        my_dic[actf]['end_date'].append(end_data)
        label = orig[index] + '-' + dest[index]
        my_dic[actf]['label'].append(label)
        node = start_data.minute / 60 + start_data.hour
        my_dic[actf]['pos'].append(node)

    first_t = min(leg_etd)
    ts = (first_t- np.datetime64('1970-01-01T00:00:00Z')
              ) / np.timedelta64(1, 's')
    first_t = datetime.utcfromtimestamp(ts)
    firt = first_t.minute / 60 + first_t.hour
    last_t =  max(leg_etd)
    ts = (last_t- np.datetime64('1970-01-01T00:00:00Z')
              ) / np.timedelta64(1, 's')
    last_t = datetime.utcfromtimestamp(ts)
    last =  last_t.minute / 60 + last_t.hour
    CreateGanttChart(my_dic, firt, last)
