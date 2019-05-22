import pandas as pd
import win32com.client
import datetime
import matplotlib.pyplot as plt
import numpy as np

############################################################
# Insert the path to your Download folder here!
downloadlink = 'C:/Users/b7217/Downloads/'
############################################################

#######################Define Traded hours##################
h24 = ['H01','H02','H03','H04','H05','H06','H07','H08','H09','H10','H11','H12','H13','H14','H15','H16','H17','H18','H19','H20','H21','H22','H23','H24']
cch24 = [['H01','H02','H03','H04','H05','H06','H07','H08','H09','H10','H11','H12','H13','H14','H15','H16','H17','H18','H19','H20','H21','H22','H23','H24'],['EUR/MWh','MWh']]
#######################Define times#########################
holnap = datetime.datetime.now() + datetime.timedelta(days=1)
holnapinap = holnap.strftime("%Y-%m-%d")
hete = datetime.datetime.now() - datetime.timedelta(days=5)
hetenap = hete.strftime("%Y-%m-%d")
otenap = holnap.strftime("%d_%m_%Y")
opcomnap = holnap.strftime('%#m_%#d_%#Y')
most = datetime.datetime.now()
bspdate = most.strftime("%B")
#################################################################
################### PROCESS RAW TABLE DATA ######################

########### OTE Table ##########
ote = pd.read_excel(downloadlink+'DM_'+otenap+'_EN'+'.xls',index_col=[0], skiprows=5,nrows=24)
otedam = ote.drop(['CZ=>SK','SK=>CZ'], axis=1)
ot = otedam.set_index([h24])
ot.columns=['OTE', 'OTE Volume']

########### IBEX Table ##########
ibex = pd.read_excel(downloadlink+hetenap+'-'+holnapinap+'.xls', header=12, index_col=[0,1])
ibex.index = pd.MultiIndex.from_product(cch24, names=['Hours', 'second'])
ib = ibex.unstack()
ibexcols = [0,1,2,3,4,5,6,7,8,9,10,11]
ibdam= ib.drop(ib.columns[ibexcols], axis=1)
ibdam.columns=['IBEX', 'IBEX Volume']
#ibdam.to_excel(downloadlink+'ibexdam.xls')

########### ISOT Table ##########
isot = pd.read_excel(downloadlink+'DailyResultsDM_'+holnapinap+'.xls', index_col=[0,1])
skcols = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15]
skdam = isot.drop(['Unsuccessful supply - Amount (MWh)','Unsuccessful demand - Amount (MWh)','Successful demand - Count','Difference - Amount (MWh)','Total supply - Amount (MWh)','Total demand - Amount (MWh)','Total supply - Count','Total demand - Count','Unsuccessful demand - Count','Successful supply - Count','Unsuccessful supply - Count'], axis=1)
skdam['ISOT Volume'] = skdam['Successful supply - Amount (MWh)'] + skdam['Successful demand - Amount (MWh)']
isotfinal = skdam.drop(['Successful supply - Amount (MWh)','Successful demand - Amount (MWh)'], axis=1)
isotdam = isotfinal.set_index([h24])
isotdam.columns=['ISOT', 'SK Volume']

########### HUPX Table ##########
hupx = pd.read_excel(downloadlink+'dam_weekly_data_export.xlsx', sheet_name='Órás árak és mennyiségek', header=3)
hupx.index = pd.MultiIndex.from_product(cch24, names=['Hours', 'second'])
hx = hupx.unstack()
hcols = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15]
hxdam= hx.drop(hx.columns[hcols], axis=1)
hxdam.columns=['HUPX', 'HUPX Volume']

########### OPCOM ##############
opcom = pd.read_csv(downloadlink+'rezultatePZU_'+opcomnap+'.csv', sep=',', skiprows=8)
rodam = opcom.drop(['Trading Zone','Interval','Traded Buy Volume [MWh]','Traded Sell Volume [MWh]'], axis=1)
opcomdam = rodam.set_index([h24])
opcomdam.columns=['OPCOM', 'OPCOM Volume']
#opcomdam.to_excel(downloadlink+'romanarak.xlsx')

########### CROPEX Table ##########
cropexda = pd.read_excel(downloadlink+'CropexDA.xls', index_col=[0], skiprows=2, nrows=24)
crcols = [0,1,2,3,4,5,6,7,8]
crdam = cropexda.drop(cropexda.columns[crcols], axis=1)
crdam.columns=['CROPEX', 'CROPEX Volume']
cropex = crdam.set_index([h24])


########### EPEX DE Table ##########
epex_de = pd.read_excel(downloadlink+'EPEXDE.xls', header=0)
epex_de.index = pd.MultiIndex.from_product(cch24, names=['Hours', 'second'])
epexde = epex_de.unstack()
decols = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17]
epexdedam= epexde.drop(epexde.columns[decols], axis=1)
epexdedam.columns=['EPEX DE', 'DE Volume']
#epexdedam.to_excel(downloadlink+'nemet.xlsx')

########### EPEX FR Table ##########
epex_fr = pd.read_excel(downloadlink+'EPEXFR.xls', header=0)
epex_fr.index = pd.MultiIndex.from_product(cch24, names=['Hours', 'second'])
epexfr = epex_fr.unstack()
frcols = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17]
epexfrdam= epexfr.drop(epexfr.columns[frcols], axis=1)
epexfrdam.columns=['EPEX FR', 'FR Volume']
#epexfrdam.to_excel(downloadlink+'francia.xlsx')


########### EPEX AT Table ##########
epex_at = pd.read_excel(downloadlink+'EPEXAT.xls', header=0)
epex_at.index = pd.MultiIndex.from_product(cch24, names=['Hours', 'second'])
epexat = epex_at.unstack()
atcols = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17]
epexatdam= epexat.drop(epexat.columns[atcols], axis=1)
epexatdam.columns=['EPEX AT', 'AT Volume']
#epexatdam.to_excel(downloadlink+'osztrak.xlsx')


########### EPEX SR Table ##########
epex_sr = pd.read_excel(downloadlink+'seepex.xls', header=0)
epex_sr.index = pd.MultiIndex.from_product(cch24, names=['Hours', 'second'])
epexsr = epex_sr.unstack()
srcols = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17]
epexsrdam= epexsr.drop(epexsr.columns[srcols], axis=1)
epexsrdam.columns=['SEEPEX', 'SR Volume']
#epexsrdam.to_excel(downloadlink+'szerb.xlsx')


########### BSP Table ##########
bspp = pd.read_excel(downloadlink+'MarketResultsAuction.xlsx', sheet_name=bspdate, index_col=[0], header=1, nrows=30)
bsppda = bspp.loc[holnapinap]
bspv = pd.read_excel(downloadlink+'MarketResultsAuction.xlsx', sheet_name=bspdate, index_col=[0], header=1, skiprows=35)
bspvda = bspv.loc[holnapinap]
########### BSP price ##########
bspprice = pd.DataFrame(data=bsppda)
bsp = bspprice.reset_index()
bsp1 = bsp.drop('index', axis=1)
bsp2 = bsp1.drop([0,1])
bsp3 = bsp2.reset_index().drop('index', axis=1).set_index([h24])
########### BSP volume ##########
bspvol = pd.DataFrame(data=bspvda)
bsv = bspvol.reset_index()
bsv1 = bsv.drop('index', axis=1)
bsv2 = bsv1.drop([0,1])
bsv3 = bsv2.reset_index().drop('index', axis=1).set_index([h24])
########### BSP data ############
bsptable = pd.concat([bsp3,bsv3], axis=1)
bsptable.columns=['BSP', 'BSP Volume']


########## CONCAT MAGIC ############

table = pd.concat([ot,ibdam,isotdam,hxdam,opcomdam,cropex,epexdedam,epexfrdam,epexsrdam,bsptable,epexatdam], axis=1)


####################################
####################################
####################################
#TO CHECK THE RAW TABLE:
#table.to_excel(downloadlink+'damtabla.xls')
####################################
####################################
####################################


########### CHART ##################
pricestable = table[['OTE','IBEX','HUPX','ISOT','OPCOM','EPEX DE','EPEX FR','EPEX AT','CROPEX','SEEPEX','BSP']]
plt.plot(pricestable)
########### LEGEND SETTINGS ########
plt.legend(pricestable, loc='upper center', bbox_to_anchor=(0.5, -0.12), ncol=6)
plt.xticks(np.arange(0,24, step=3))
plt.xlabel('Hours')
plt.ylabel('Prices')
plt.title('CWE & CEE Hourly Prices')
plt.gca().yaxis.grid(True)
#plt.show()
#CREATE FIG JPG
plt.savefig(downloadlink+'DAM.jpg',dpi=130,bbox_inches = "tight")
#CREATE FIG JPG

########## FINAL TABLE #############
table.loc['Baseload (€/MWh)'] = table[:24].mean()
table.loc['Peakload (€/MWh)'] = table[8:20].mean()
table.loc['Volume (MWh)'] = table[:24].sum()
pr = table[['OTE','IBEX','HUPX','ISOT','OPCOM','EPEX DE','EPEX FR','EPEX AT','CROPEX','SEEPEX','BSP']].iloc[-3:-1]
vol = table[['OTE Volume','IBEX Volume','HUPX Volume','SK Volume','OPCOM Volume','DE Volume','FR Volume','AT Volume','CROPEX Volume','SR Volume','BSP Volume']].iloc[-1:]
vol.columns = ['OTE','IBEX','HUPX','ISOT','OPCOM','EPEX DE','EPEX FR','EPEX AT','CROPEX','SEEPEX','BSP']
#Merge datasets
summary = pd.concat([pr,vol]).round(2)
del summary.index.name
#.apply('€{:.2f}'.format)
#formatters = {'Baseload':"€{x:.2f}".format}

########## STYLE TABLE ############
html = (
    summary.style
    .format('{:.2f}')
    .set_table_styles([{'selector': 'th', 'props': [('font-size', '12pt'),('text-align', 'center'),('font-family', 'Calibri'),('padding', '2pt')]}])
    .set_properties(**{'font-size': '11pt', 'font-family': 'Calibri','border-collapse': 'collapse','border': '1px solid black','text-align':'center', 'padding':'5pt'})
    .render()
)


###########################################################################
kep = downloadlink+'DAM.jpg'
now = datetime.datetime.now() + datetime.timedelta(days=1)
date = now.strftime("%Y-%m-%d")
#some constants (from http://msdn.microsoft.com/en-us/library/office/aa219371%28v=office.11%29.aspx)
olFormatHTML = 2
olFormatPlain = 1
olFormatRichText = 3
olFormatUnspecified = 0
olMailItem = 0x0
############################ CREATE MAIL ##################################
obj = win32com.client.Dispatch("Outlook.Application")
newMail = obj.CreateItem(olMailItem)
newMail.Subject = "Napi tőzsdeárak "+date
newMail.BodyFormat = olFormatHTML    #or olFormatRichText or olFormatPlain
attachment = newMail.Attachments.Add(kep)
attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId1")
newMail.HTMLBody = html+"<tr>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</tr><html><body><img src=""cid:MyId1""></body></html>"
newMail.To = "EMAIL ADDRESS COMES HERE"

# carbon copies and attachments (optional)

#newMail.CC = "moreaddresses here"
#newMail.BCC = "address"
#attachment1 = "Path to attachment no. 1"
#attachment2 = "Path to attachment no. 2"
#newMail.Attachments.Add(kep)
#newMail.Attachments.Add(attachment2)
# open up in a new window and allow review before send
newMail.display()

