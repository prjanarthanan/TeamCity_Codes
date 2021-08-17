from pandas import read_excel
from numpy import array
from time import sleep
import getpass
from tkinter import filedialog
import matplotlib.pyplot as plt

def read_file(filename):
    dataset = read_excel(filename)
    y = dataset.groupby('Assignment group').count()
    for group in list(y.index):
        x = dataset.loc[:,['Assignment group','Number','Short description']]
        x = x[x['Assignment group'] ==group]
        #seg_group(i,dataset)
        seg_group(x,group)
def seg_group(dataset,group):
    Group = group
    x=dataset
    c = ['Job completed abnormally','Job not ready by end of its time window','Job failed','Job deferred',
     'SecureTransport - Route Execution Failed','Agent unavailable for job','Error occurred while launching job']
    k = []
    for description in x.iloc[:,2]:
        o = [val for val in c[:] if val.lower() in description.lower()]
        if o == []:
            o = 'Others'
        if o == 'Others':
            k.append(o)
        else:
            k.append(str(o[-1]))
    x['status'] = array(k)
    x = x.iloc[:,[0,3]]
    print("CLassifying groups are completed")
    sleep(1)
    ak = x.groupby(by='status').count()
    login_user = getpass.getuser()
    outfile = 'C:\\Users\\'+login_user+'\\Analysis_'+Group+'.xlsx'
    pie_out = 'C:\\Users\\'+login_user+'\\Pie_chart_'+Group+'.png'
    ak.to_excel(outfile,sheet_name=Group)
    labels = x.groupby(by='status').groups.keys()
    plt.figure(figsize=(10,8))
    plt.pie(ak['Assignment group'],labels=labels)
    plt.title('Batch team  analysis')
    plt.legend(loc='upper right')
    plt.savefig(pie_out)
    print(f"Activity completed for {Group}, check the output file and pie chart")
    sleep(1)

print("Select the excel file to be imported")

input_file = filedialog.askopenfile(mode="r", title="Select File", 
        filetypes=[("All Files", "*.xlsx")])

print("File import completed")

sleep(1)

filename = input_file.name

read_file(filename)
