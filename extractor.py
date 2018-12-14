ls_id=['Vehicle definition','AIVC Activation status', 'USB Definition', 'Ecall Zone', 'Car Variant', 'DataRecord', 'Vehicle Manufacturer Spare Part R', 'ConfigurationFileReferenceLink data', 'VehicleManuECUHWNumber', 'System Supplier ECU SW', 'HWVariantID', 'TCU state','IMEI', 'CCID', 'Identificatio (IMSI)', 'eUICC EID','eUICC MSISDN', 'MQTT payload compression']

data=[]

def main():
	with open("DiagConsoleLog.txt", "r") as f:
		l = []
		for line in f : 
			line = line.strip()
			l = line.split('\t') 
			#print(l)
			for elem in ls_id :
				for k in l :
					if elem in k :
						try : 
							print(elem + ' =  '+l[l.index(k)+1] )
							ls_id.remove(elem)
							data.append((elem ,l[l.index(k)+1]))
						except IndexError : 
							print(l[l.index(k)])
	#print(data)
	
if __name__ == "__main__" : 
	main()
