import xml.etree.ElementTree as ET
import csv
import tkinter
from functools import partial
from tkinter.filedialog import askopenfilename, asksaveasfilename
import pandas as pd
import xml.etree.ElementTree as et

# # NEROZBEHANE -> PRILIS KOMPLIKOVANE XML - https://medium.com/@robertopreste/from-xml-to-pandas-dataframes-9292980b1c1c
# # PRIKLAD POUZITIA - df_tois = parse_XML('TOIS_test.xml', ['key', 'summary'])
def parse_XML(xml_file, df_cols):
    """Parse the input XML file and store the result in a pandas
    DataFrame with the given columns.

    The first element of df_cols is supposed to be the identifier
    variable, which is an attribute of each node element in the
    XML data; other features will be parsed from the text content
    of each sub-element.
    """

    xtree = et.parse(xml_file)
    xroot = xtree.getroot()
    rows = []

    for node in xroot:
        res = []
        res.append(node.attrib.get(df_cols[0]))
        for el in df_cols[1:]:
            if node is not None and node.find(el) is not None:
                res.append(node.find(el).text)
            else:
                res.append(None)
        rows.append({df_cols[i]: res[i]
                     for i, _ in enumerate(df_cols)})

    out_df = pd.DataFrame(rows, columns=df_cols)

    return out_df


# mozny manualny vyber suboru
tkinter.Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
filename = askopenfilename(initialdir="/", title="Nacitat subor", filetypes=(
    ("xml files", "*.xml"), ("all files", "*.*")))  # show an "Open" dialog box and return the path to the selected file

tree = ET.parse(filename)
root = tree.getroot()

# prvy incident
# rss/channel/item struktura
# print(root[1][6][0].text)

# premenne, ktore sa budu nacitavat do riadku
tois_im = ''
tois_key = ''
tois_title = ''
tois_prio = ''
tois_label = []
tois_state = ''
tois_created = ''
tois_updated = ''
tois_link = ''
n_item = 0  # kontrola cyklu

# kde sa ma csv ulozit
filename = asksaveasfilename(initialdir="/", title="Ulozit ako", filetypes=(
    ("csv files", "*.csv"), ("all files", "*.*")))  # show an "Open" dialog box and return the path to the selected file

# vytvorenie vystupneho csv
tois_output = open(filename, encoding="utf-8", mode='w+') # predpoklada sa xml ako pripona
csvwriter = csv.writer(tois_output, delimiter=";", lineterminator='\n')
headers = ["IM", "key", "title", "prio", "label", "state", "created", "updated", "link"]
csvwriter.writerow(headers)

# vytvorenie txt kontrolenho vystupu
tois_txt_out = open(filename + '.aux', encoding="utf-8", mode='w+')
tois_txt_out.write("IM; key; title; prio; label; state; created; updated; link\n")

# data z xml, encode kvoli diakritike -> vypisovat tu mozne jedine cez .decode('utf-8') prikaz na konci
for item in root.findall('./channel/item'):
    tois_key = item.find('key').text
    if item.find('summary').text[:2] == "IM":
        tois_im = item.find('summary').text.split(' ', 1)[0]  # len prve slovo -> IM cislo
        tois_title = item.find('summary').text[len(tois_im) + 3:]
    else:
        for customname in item.findall("./customfields/customfield[@id='customfield_11343']"):
            tois_im = customname[1][0].text
        tois_title = item.find('summary').text

    tois_prio = item.find('priority').text
    if tois_prio == 'Critical':
        tois_prio = '1'
    elif tois_prio == 'Major':
        tois_prio = '2'
    elif tois_prio == 'Minor':
        tois_prio = '3'
    else:
        tois_prio = '0'
    tois_label = []
    for label in item.find('labels'):  # moze byt aj viac labelov teoreticky
        tois_label.append(label.text)
    if not tois_label:
        tois_label_str = 'L2'
    elif len(tois_label) == 1:
        tois_label_str = tois_label[0]
    else:
        tois_label_str = ', '.join(map(str, tois_label))
    tois_state = item.find('status').text
    tois_created = item.find('created').text  # .split(' ', 1)[1].split(' +', 1)[0]  # len datum bez dna a casovej zony
    tois_updated = item.find('updated').text  # .split(' ', 1)[1].split(' +', 1)[0]  # len datum bez dna a casovej zony
    tois_link = item.find('link').text

    # zapis v tomto poradi do riadku
    csvwriter.writerow([tois_im, tois_key, tois_title, tois_prio,
                        tois_label_str, tois_state, tois_created, tois_updated, tois_link])

    # zapis to txt
    tois_txt_out.write(tois_im + '; ' + tois_key + '; ' + tois_title + '; ' + tois_prio + '; ' + tois_label_str + '; '
                       + tois_state + '; ' + tois_created + '; ' + tois_updated + '; ' + tois_link + '\n')

    n_item = n_item + 1

tois_output.close()
tois_txt_out.close()
print('Konverzia do csv uspesna - spracovanych ' + str(n_item) + ' incidentov!')


# # open a file for writing
#
# tois_data = open('/tmp/ResidentData.csv', 'w')
#
# # create the csv writer object
#
# csvwriter = csv.writer(tois_data)
# tois_head = []

# count = 0
# for member in root.findall('item'):
# 	resident = []
# 	address_list = []
# 	if count == 0:
# 		name = member.find('Name').tag
# 		resident_head.append(name)
# 		PhoneNumber = member.find('PhoneNumber').tag
# 		resident_head.append(PhoneNumber)
# 		EmailAddress = member.find('EmailAddress').tag
# 		resident_head.append(EmailAddress)
# 		Address = member[3].tag
# 		resident_head.append(Address)
# 		csvwriter.writerow(resident_head)
# 		count = count + 1
#
# 	name = member.find('Name').text
# 	resident.append(name)
# 	PhoneNumber = member.find('PhoneNumber').text
# 	resident.append(PhoneNumber)
# 	EmailAddress = member.find('EmailAddress').text
# 	resident.append(EmailAddress)
# 	Address = member[3][0].text
# 	address_list.append(Address)
# 	City = member[3][1].text
# 	address_list.append(City)
# 	StateCode = member[3][2].text
# 	address_list.append(StateCode)
# 	PostalCode = member[3][3].text
# 	address_list.append(PostalCode)
# 	resident.append(address_list)
# 	csvwriter.writerow(resident)
# Resident_data.close()
