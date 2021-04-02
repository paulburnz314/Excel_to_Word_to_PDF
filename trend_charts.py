import openpyxl
from hard_data import SIU_LIST
from output.list_dict_surcharges_21_output import monthly_surcharges


if __name__ == '__main__':
	month = "January_2021"
	row_num = 4
	workbook = openpyxl.load_workbook("./input/surcharge_chart_data_add_to.xlsx")
	sheet_flows = workbook.worksheets[0]
	sheet_tss = workbook.worksheets[1]
	sheet_cbod = workbook.worksheets[2]
	sheet_nh3n = workbook.worksheets[3]
	sheet_tphos = workbook.worksheets[4]
	temp_data = None
	temp_flows = []
	temp_cbod = []
	temp_tss = []
	temp_nh3n = []
	temp_tp = []

	for user in SIU_LIST:
		for sub_dict in monthly_surcharges:
			if sub_dict["Month_Year"] == month and sub_dict["User_Code"] == user:
				temp_data = sub_dict

		flow = float(temp_data["Flow"])
		temp_flows.append(flow)

		cbod_ppm = float(temp_data["CBOD_ppm"])
		cbod_load_total = round((flow * cbod_ppm * 8.34), 1)
		cbod_load_charged = float(temp_data["CBOD_load"])
		cbod_amt_charged = float(temp_data["CBOD_amt"])
		temp_cbod.append(cbod_load_total)
		temp_cbod.append(cbod_load_charged)
		temp_cbod.append(cbod_amt_charged)
		temp_cbod.append(cbod_ppm)

		tss_ppm = float(temp_data["TSS_ppm"])
		tss_load_total = round((flow * tss_ppm * 8.34), 1)
		tss_load_charged = float(temp_data["TSS_load"])
		tss_amt_charged = float(temp_data["TSS_amt"])
		temp_tss.append(tss_load_total)
		temp_tss.append(tss_load_charged)
		temp_tss.append(tss_amt_charged)
		temp_tss.append(tss_ppm)

		nh3n_ppm = float(temp_data["NH3N_ppm"])
		nh3n_load_total = round((flow * nh3n_ppm * 8.34), 1)
		nh3n_load_charged = float(temp_data["NH3N_load"])
		nh3n_amt_charged = float(temp_data["NH3N_amt"])
		temp_nh3n.append(nh3n_load_total)
		temp_nh3n.append(nh3n_load_charged)
		temp_nh3n.append(nh3n_amt_charged)
		temp_nh3n.append(nh3n_ppm)

		tp_ppm = float(temp_data["TP_ppm"])
		tp_load_total = round((flow * tp_ppm * 8.34), 1)
		tp_load_charged = float(temp_data["TP_load"])
		tp_amt_charged = float(temp_data["TP_amt"])
		temp_tp.append(tp_load_total)
		temp_tp.append(tp_load_charged)
		temp_tp.append(tp_amt_charged)
		temp_tp.append(tp_ppm)

	month = month.replace("_", " ")
	workbook.active = sheet_flows
	temp_flows.insert(0, month)
	sheet_flows.append(temp_flows)
	temp_tss.insert(0, month)
	sheet_tss.append(temp_tss)
	temp_cbod.insert(0, month)
	sheet_cbod.append(temp_cbod)
	temp_tp.insert(0, month)
	sheet_tphos.append(temp_tp)
	temp_nh3n.insert(0, month)
	sheet_nh3n.append(temp_nh3n)

	workbook.save(filename='./output/surcharge_chart_data_output_0221z.xlsx')
