import argparse
import json
import logging
import time
from typing import Any

import openpyxl
import pandas
from openpyxl.utils import get_column_letter

logging.basicConfig(format="%(asctime)s | %(levelname)s | %(message)s", level=logging.INFO, datefmt="%Y-%m-%d %H:%M:%S")


def append_to_excel(excel_path: str, data_frame: pandas.DataFrame, sheet_name: str, full_name: str):
	with pandas.ExcelWriter(excel_path, mode="a", engine="openpyxl") as excel_file:
		data_frame.to_excel(excel_file, sheet_name=sheet_name, startcol=2, startrow=0)
	change_cell(excel_path, full_name, sheet_name)


def change_cell(excel_path, full_name, sheet_name):
	excel = openpyxl.open(excel_path)
	active_sheet = excel[sheet_name]
	active_sheet['A1'] = full_name
	excel.save(excel_path)
	excel.close()


def convert_json_to_excel(input_file: str, output_file: str):
	json_string = extract_json(input_file)
	json_df = pandas.json_normalize(json_string)
	create_workbook(output_file)
	tables = extract_dataframes(json_df)
	for name, table in sorted(tables.items()):
		sheet_name = create_short_name(name)
		append_to_excel(output_file, table, sheet_name, name)
	format_excel(output_file)


def create_short_name(name: str) -> str:
	names = name.split(".")
	short_name = ""
	is_to_long = len(name) > 31
	if is_to_long:
		for value in names[:len(names) - 1]:
			short_name = f"{short_name}{value[0:2]}{value[len(value) - 1]}."
		potential_short_name = f"{short_name}{names[len(names) - 1]}"
		if len(potential_short_name) <= 31:
			short_name = potential_short_name
		else:
			short_name = f"{potential_short_name[:30]}{potential_short_name[len(potential_short_name) - 1]}"
	else:
		short_name = name
	return short_name


def create_workbook(output_file: str):
	workbook = openpyxl.Workbook()
	sheet = workbook.active
	sheet.title = "temp"
	workbook.save(filename=output_file)


def extract_dataframes(df):
	columns_list = []
	new_tables = {}
	for field, value in df.iteritems():
		if not isinstance(value.values[0], list):
			columns_list.append(value.name)
		else:
			new_tables[value.name] = pandas.json_normalize(value.values[0])
	new_tables["ROOT"] = df[columns_list]

	loop_again = False
	while True:
		for name, table in new_tables.copy().items():
			changed_table = new_tables[name]
			for column, cells in table.iteritems():
				i = 1
				change_table = False
				for value in cells.values:
					if isinstance(value, list):
						new_column = f"{name}.{column}{i}"
						new_df = pandas.json_normalize(value)
						new_tables[new_column] = new_df
						change_table = True
					i += 1
				if change_table:
					print(f"Table: {name}, column count before: {len(table.columns)}")
					new_tables.pop(name)
					new_tables[name] = table.drop(column, axis=1)
					new_table = new_tables[name]
					print(f"Table: {name}, column count after: {len(new_table.columns)}")
					loop_again = True

		if not loop_again:
			break
		else:
			loop_again = False
	return new_tables


def extract_json(input_file: str) -> Any:
	with open(input_file, encoding="utf-8") as json_file:
		json_data = json.load(json_file)
	return json_data


def format_excel(output_file: str):
	excel = openpyxl.open(output_file)
	excel.remove(excel["temp"])
	sheets = excel.sheetnames
	for sheet in sheets:
		active_sheet = excel[sheet]
		active_sheet.sheet_view.showGridLines = False
		active_sheet.freeze_panes = 'A2'
		active_sheet['c1'] = "nr"
		for column in active_sheet.columns:
			column_name = get_column_letter(column[0].column)
			maximum_value = 0
			for cell in active_sheet[column_name]:
				val_to_check = len(str(cell.value))
				if val_to_check > maximum_value:
					maximum_value = val_to_check
			active_sheet.column_dimensions[column_name].width = maximum_value + 1
	excel.save(output_file)
	excel.close()


def parse_arguments() -> dict[str, Any]:
	argument_parser = argparse.ArgumentParser()
	argument_parser.add_argument(
		"-i", "--inputpath", required=True, help="Path to the input file"
	)
	argument_parser.add_argument(
		"-o", "--outputpath", required=True, help="Path to the output file",
	)
	cli_arguments = vars(argument_parser.parse_args())
	return cli_arguments


if __name__ == '__main__':
	arguments = parse_arguments()
	input_path = arguments["inputpath"]
	output_path = arguments["outputpath"]

	start_time = time.perf_counter()
	convert_json_to_excel(input_path, output_path)
	end_time = time.perf_counter()

	logging.info(f"Excel file with tables created in {end_time - start_time:0.4f} seconds: {output_path}")
