import xlrd
from xlutils.copy import copy as xlcopy
import xml.etree.ElementTree as ET
from xml.etree.ElementTree import ParseError

INPUT_FILE = 'TestReport_general.xlsx'
OUTPUT_FILE = 'SmoothReport.xls'

def clear_step(step):
    """ Removes specifed tags from steps text """
    tags = (
    '<DIV>',
    '<P>',
    '</DIV>',
    '</P>',
    '<P />'
    )

    if step:
        for tag in tags:
            step = step.replace(tag, '')
    else:
        # if step is None then just return emtpy string
        step = ''

    return step

def parse_steps(xmldata):
    """
    Parse xml data from step description and returns two lists with
    steps and expected results
    """
    try:
        xml = ET.fromstring(xmldata)
    except ParseError:
        # if we cannot parse data as XML then just return this data as is
        return xmldata

    result = [[], []]
    # some python magic
    for step in xml.iter('step'):
        for ps, dest in zip(step.iter('parameterizedString'), result):
            dest.append(ps.text)
    steps_raw, expected_results_raw = result

    # clear steps and expected results from html formatting
    steps = [ clear_step(s) for s in steps_raw ]
    expected_results = [ clear_step(r) for r in expected_results_raw ]

    return (steps, expected_results)


# main logic
def main():
    # open xls to process
    workbook = xlrd.open_workbook(INPUT_FILE)
    sheet = workbook.sheet_by_index(0)

    # get writable copy
    result_book = xlcopy(workbook)
    result_sheet = result_book.get_sheet(0)
    for row_i in range(1, sheet.nrows):
        steps = None
        expected_results = None
        for col_i in range(sheet.ncols):
            value = sheet.cell_value(row_i, col_i)

            # column is Action
            if col_i == 2:
                try:
                    steps, expected_results = parse_steps(value[2:])
                except ValueError:
                    pass
                value = "\n".join(steps) if steps else value

            # column is Expected Result
            if col_i == 3 and expected_results:
                value = "\n".join(expected_results)
                expected_results = None

            result_sheet.write(row_i, col_i, value)

    result_book.save(OUTPUT_FILE)


if __name__ == '__main__':
    main()
