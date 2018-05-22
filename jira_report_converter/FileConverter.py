#!/usr/bin/env python3

import argparse

from ExcelConverter import ExcelConverter

if __name__ == '__main__':
	parser = argparse.ArgumentParser()
	parser.add_argument("-f",  "--file", type=str, required=True, help="input file '.xls'")
	args = parser.parse_args()

	print(args.file)

	fIn = open(args.file, 'rb')
	converter = ExcelConverter(fIn)
	fIn.close()
	outBytes = converter.convert()
	outBytes.seek(0)

	file_name = converter.get_file_name()
	fOut = open(file_name, 'wb')
	fOut.write(outBytes.read())
	outBytes.close()
	fOut.close()
