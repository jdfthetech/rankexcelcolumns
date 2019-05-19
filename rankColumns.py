#!/usr/bin/python3


import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

def main():
	# import the excel file (name and sheet name must be exact)

	df = pd.read_excel('importTest.xlsx', sheet_name='Sheet1')

	# create Rank column, use the algorithm to populate the rank number

	df['Rank'] = df['Spend'] / df['Income']

	# sort by the rank number

	df.sort_values(by=['Rank'], ascending=False, inplace=True)

	# drop some useless columns

	df = df.drop(columns=['header1','header2'])

	# write results file

	df.to_excel('results.xlsx', index=False)

if __name__ == '__main__':
	main()


