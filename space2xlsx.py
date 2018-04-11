#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import re
import os
import argparse
import xlsxwriter

def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        pass

#    try:
#        import unicodedata
#        unicodedata.numeric(s)
#        return True
#    except (TypeError, ValueError):
#        pass

    return False

def convert_space(lines, boolInter):
    new_lines = list()

    for line in lines:
        re_result, number = re.subn('\s+', '\t', line)

        if len(re_result) and re_result[0] == '\t':
            re_result = re_result[1:]
        if len(re_result) and re_result[-1] == '\t':
            re_result = re_result[0:-2]
        new_lines.append(re_result)

    while len(new_lines) > 0:
        if len(new_lines[0]) == 0:
            new_lines = new_lines[1:]
        elif not is_number(new_lines[0][0]):
            new_lines = new_lines[1:]
        elif not is_number(new_lines[0][-1]):
            new_lines = new_lines[1:]
        elif is_number(new_lines[0][0]) and is_number(new_lines[0][-1]):
            if not boolInter:
                break
            else:
                print("The first line to read: " + new_lines[0] + '\n')
                if input("Is this correct? y/n") in ['y', 'Y', '']:
                    break

    line_index = 0
    while line_index < len(new_lines) and is_number(new_lines[line_index][0]) and is_number(new_lines[line_index][-1]):
        line_index = line_index + 1
    new_lines = new_lines[0:line_index]

    output_lines = [line.split() for line in new_lines]

    return output_lines


def main():
    usage = "Usage: " + sys.argv[0] + " -a / -f input_file [-o output_file] [-inter]\n"\
          + "       -a     : Process all files in current folder. \n" \
          + "       -o     : Output File Name or Prefix. \n"\
          + "       -f     : Input File Name if '-a' flag not specified\n"\
          + "       -inter : Interactive choose first lines to read\n"\
          + "       -cbn   : Combine all files to one worksheet"

    # Check arguments
    if not (len(sys.argv) in (2, 3, 4, 5, 6)):
        print("ERROR: Wrong Number of Arguments Provided")
        print(usage)
        exit(1)

    flag_all_file = False
    flag_output_file = False
    flag_interactive = False
    flag_conbine     = False

    # Get arguments
    arg_count = 1
    while arg_count < len(sys.argv):
        if sys.argv[arg_count] == '-h':
            print(usage)
            quit()
        if sys.argv[arg_count] == '-a':
            flag_all_file = True
        elif sys.argv[arg_count] == '-o':
            output_file_name = sys.argv[arg_count+1]
            flag_output_file = True
            arg_count += 1
        elif sys.argv[arg_count] == '-f':
            if flag_all_file:
                print("Ignoring input file name because \'-a\' flag.")
            else:
                input_file_name = sys.argv[arg_count+1]
            arg_count += 1
        elif sys.argv[arg_count] == '-inter':
            flag_interactive = True
        elif sys.argv[arg_count] == '-cbn':
            flag_conbine == True
        arg_count += 1

    # build lists for input and output file

    if flag_all_file:

        input_file_names = list()

        for filename in os.listdir("./"):
            if not filename.startswith('.'):
                input_file_names.append(filename)

    else:
        if not 'input_file_name' in locals().keys():
            print("You must specify input file unless you have -a flag.\n")
            print(usage)
            quit()
        input_file_names = [input_file_name]

    if flag_output_file and output_file_name[-5:-1] == '.xlsx':
        output_file_name = output_file_name
    elif flag_output_file and output_file_name[-4:-1] == '.xls':
        output_file_name = output_file_name + 'x'
    elif flag_output_file:
        output_file_name = output_file_name + '.xlsx'
    else:
        output_file_name = 'space2xlsx_output.xlsx'

    # Read Lines

    output_file = xlsxwriter.Workbook(output_file_name)

    if flag_conbine == True:
        for input_file_name in input_file_names:
            worksheet = output_file.add_worksheet(input_file_name)
            try:
                with open(input_file_name, "r") as input_file:
                    lines = convert_space(input_file.read().splitlines(), flag_interactive)
                    for row in range(len(lines)):
                        for colomn in range(len(lines[row])):
                            worksheet.write(row, colomn, float(lines[row][colomn]))
                    print("FINISH: " + input_file_name)
            except FileNotFoundError:
                print("ERROR: Input File\"" + input_file_name + "\" not found.")
                print(usage)
                exit(1)
    else:
        worksheet = output_file.add_worksheet("Worksheet1")
        outputColumn = 0
        for input_file_name in input_file_names:
            try:
                with open(input_file_name, "r") as input_file:
                    lines = convert_space(input_file.read().splitlines(), flag_interactive)
                    for i in range(len(lines[0])):
                        worksheet.write(0, outputColumn + i, input_file_name + "_" + str(i))
                    for row in range(len(lines)):
                        for colomn in range(len(lines[row])):
                            worksheet.write(row + 1, colomn + outputColumn, float(lines[row][colomn]))
                    outputColumn = outputColumn + len(lines[0])
                    print("FINISH WITH READ: " + input_file_name)
            except FileNotFoundError:
                print("ERROR: Input File\"" + input_file_name + "\" not found.")
                print(usage)
                exit(1)

    output_file.close()


main()
