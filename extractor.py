import openpyxl as O 

checklist = 'MS-TCU-AIVC_200701201841003187-D4b-ID20_2018-12-13.xlsm'

ls_id=['Vehicle definition','AIVC Activation status', 'USB Definition', 'Ecall Zone', 'Car Variant', 'DataRecord', 'Vehicle Manufacturer Spare Part R', 'ConfigurationFileReferenceLink data',
 'VehicleManuECUHWNumber', 'System Supplier ECU SW', 'HWVariantID', 'TCU state','IMEI', 'CCID', 'Identificatio (IMSI)', 'eUICC EID','eUICC MSISDN', 'MQTT payload compression']

ls_pos=['E46', 'E47', 'E48' , 'E49', 'E50' , 'E51' , 'E53', 'E54', 'E55', 'E56', 'E57' , 'E58' , 'E59' , 'E60', 'E61', 'E62', 'E63', 'E64']

L = list(zip(ls_id, ls_pos))

data=[]

def core():
	with open("DiagConsoleLog.txt", "r") as f:
		l = []
		for line in f : 
			line = line.strip()
			l = line.split('\t') 
			for elem in L :
				for k in l :
					if elem[0] in k :
						try : 
							print(elem[0] + ' =  '+l[l.index(k)+1] )
							L.remove(elem)
							data.append((elem[0], elem[1] ,l[l.index(k)+1]))
						except IndexError : 
							print(l[l.index(k)])

def fill_checklist(path, data_ls): 
	ww = O.load_workbook(filename=path)
	ws = ww.worksheets[2]
	for elem in data_ls: 
		ws[elem[1]] = elem[0]

if __name__ == "__main__" : 
	core()
	fill_checklist(checklist, data)
