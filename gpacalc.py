import pandas as pd
import argparse
from openpyxl import load_workbook

"""
High school GPA Calculator. Reads .csv or .xlsx file and parses class levels from spreadsheet labels to apply appropriate multipliers.
Spreasheet should contain classes as rows and grading periods as columns, and class level indicators (ex. 'AP') should be a separate word in row labels.
Functionality for 100 or 4.0 scale, custom class levels and multipliers, and saving GPAs to spreadsheet.
Defaults to 100-point scale and multipliers of 1.2 for 'AP' and 1.1 for 'Pre-AP'
"""


def arg_parse():

    """
    Parse command line arguments and return argparse object.

    args:
        None

    returns:
        args: argparse object with parsed commandline arguments
    """


    parser = argparse.ArgumentParser()
    parser.add_argument('file',type = str,help = 'excel or csv file containing grades, including extension')
    parser.add_argument('-m','--multi',type = str, metavar = 'course level  multiplier',nargs = '*',help = 'custom course levels and multipliers (case sensitive), ex: IB 1.2 Honors 1.1')
    parser.add_argument('-f','--four',action = 'store_true', help = 'generates gpa on a 4.0 scale')
    parser.add_argument('-p','--printg',action = 'store_true',help = 'prints gradebook and multiplier dictionary')
    parser.add_argument('-s','--save',action = 'store_true',help = 'appends gpa\'s to gradebook spreadsheet (creates new sheet for .xlsx files)')
    return parser.parse_args()

def open_sheet(file_name):

    """
    Open spreadsheet and put data into a DataFrame.

    args:
        file_name: string containing spreadsheet file name

    returns:
        pandas DataFrame of gradebook
    """

    try:
        # check to see if file is .xlsx, if so, put into DataFrame
        if (file_name.endswith('.xlsx')):
            grades = pd.read_excel(file_name)
            return grades
        # check to see if file is.csv, if so, put into DataFrame
        elif (file_name.endswith('.csv')):
            grades = pd.read_csv(file_name,index_col = 0)
            return grades
        # file doesn't end in either .xlsx or .csv extension, raise error
        else:
            raise argparse.ArgumentTypeError('File must include extension (either .csv or .xlsx)')
    except Exception as ex:
        print('Error in opening file:',ex)
        exit()

def parse_multipliers(multi_strings):

    """
    Parse the list of strings generated from multiplier argument into dictionary of class level keys and multiplier values.

    args:
        multi_strings: list of strings parsed from command line, ordered by label then multiplier (ex. ['AP', '1.2', 'Pre-AP', '1.1')

    returns:
        dictionary of string keys and float values, representing class level label and multiplier value
    """

    try:
        # gets labels from every other string, starting at [0]
        labels = multi_strings[::2]
        # gets multipliers from every other string, starting at [0]
        multipliers = multi_strings[1::2]
        multipliers = [float(x) for x in multipliers]
        # checks to see if there is an equal number of labels and multipliers
        if len(multipliers) != len(labels):
            raise argparser.ArgumentTypeError
        # puts lists into a dictionary
        return dict(zip(labels,multipliers))
    except Exception as ex:
        print('Error in class labels and multipliers. Make sure you have corresponding labels and multipliers, respectively.', ex)
        exit()

def add_multipliers(grades, multiplier_dict):

    """
    Multiply grades in DataFrame class-wise, according to multiplier interpreted from class level, and produced new DataFrame.

    args:
        grades: DataFrame containing gradebook
        multiplier_dict: dictionary of class label keys and multiplier values

    returns:
        second DataFrame, containing the gradebook with weighted grades
    """

    weighted_grades = grades.copy()
    # iterate through class level labels in dictionary
    for label in multiplier_dict.keys():
        # match class label in DataFrame index name to label in dictionary
        # uses regex to find label as a white-space separated word (so 'Pre-AP' is not identified as 'AP', and so forth)
        # resulting boolean mask is then reapplied to DataFrame and the filtered rows are multiplied by the appropriate multiplier
        # this is a complicated line of code :)
        weighted_grades[weighted_grades.index.str.match('^(.*\\s+)?'+label+'(\\s+.*)?$')] *= multiplier_dict[label]
    return weighted_grades

def write_to_sheet(gpa_df, grades, file_name):

    """
    Write weighted and unweighted GPA values to spreadsheet file.
    For .csv files, append GPAs as two new rows at the bottom.
    For .xlsx files, create new sheet containing only the GPAs.
    """

    try:
        # is file .xlsx? if so, use openpyxl to write GPAs to new sheet
        if (file_name.endswith('.xlsx')):
            with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
                    writer.book = load_workbook(file_name)
                    gpa_df.to_excel(writer,sheet_name = 'GPA',header=False)
        # is file .csv? if so, append GPAs to bottom of sheet as two rows
        elif (file_name.endswith('.csv')):
            gpa_df.to_csv(file_name, header = False, mode = 'a')
        print('Appended GPA\'s to', file_name)
    except Exception as ex:
        print('Error in appending GPA to gradebook:', ex)

def main():

    """
    Main method, execute GPA calculation with specified arguments.
    """
    # parse commandline arguments
    args = arg_parse()

    # open gradebook as DataFrame
    grades = open_sheet(args.file)

    # set default mutliplier dictionary
    multiplier_dict = {'AP':1.2,'Pre-AP':1.1}

    # if custom multipliers are specified, make them into a dictionary
    if args.multi is not None:
        multiplier_dict = parse_multipliers(args.multi)
    # make another gradebook and add multipliers to grades
    weighted_grades = add_multipliers(grades,multiplier_dict)


    # calculate the GPAs!
    gpa = grades.mean().mean()
    weighted_gpa = weighted_grades.mean().mean()

    # if a 4.0 GPA was specified, divide GPAs by 25
    if args.four:
        weighted_gpa = round(weighted_gpa/25,1)
        gpa = round(gpa/25,1)
    else:
        weighted_gpa = round(weighted_gpa,1)
        gpa = round(gpa,1)

    # if asked to print multiplier assignments and gradebook, do it!
    if args.printg:
        print('Multiplier assigments:',multiplier_dict)
        print(grades)
        print(weighted_grades)

    # print the GPAs (finally)
    print('Unweighted GPA:',gpa)
    print('Weighted GPA:',weighted_gpa)

    # if asked to save GPAs back to gradebook, do it!
    if args.save:
        #create a 2x1 data frame containing GPAs
        gpa_df = pd.DataFrame()
        gpa_df.loc['Unweighted GPA',grades.columns[0]] = gpa
        gpa_df.loc['Weighted GPA',grades.columns[0]] = weighted_gpa
        #append GPAs to gradebook file
        write_to_sheet(gpa_df,grades,args.file)


if __name__ == '__main__':
    main()
