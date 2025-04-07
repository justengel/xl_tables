import os
import win32com.client  # Requires "pip install pywin32"


__all__ = ['get_xls_properties', 'get_file_details']


# https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.tools.excel.workbook.builtindocumentproperties?view=vsto-2017
BUILTIN_XLS_ATTRS = ['Title', 'Subject', 'Author', 'Keywords', 'Comments', 'Template', 'Last Author', 'Revision Number',
                     'Application Name', 'Last Print Date', 'Creation Date', 'Last Save Time', 'Total Editing Time',
                     'Number of Pages', 'Number of Words', 'Number of Characters', 'Security', 'Category', 'Format',
                     'Manager', 'Company', 'Number of Btyes', 'Number of Lines', 'Number of Paragraphs',
                     'Number of Slides', 'Number of Notes', 'Number of Hidden Slides', 'Number of Multimedia Clips',
                     'Hyperlink Base', 'Number of Characters (with spaces)']


def get_xls_properties(filename, xl=None):
    """Return the known XLS file attributes for the given .xls filename."""
    quit = False
    if xl is None:
        xl = win32com.client.DispatchEx('Excel.Application')
        quit = True

    # Open the workbook
    wb = xl.Workbooks.Open(filename)

    # Save the attributes in a dictionary
    attrs = {}
    for attrname in BUILTIN_XLS_ATTRS:
        try:
            val = wb.BuiltinDocumentProperties(attrname).Value
            if val:
                attrs[attrname] = val
        except:
            pass

    # Quit the excel application
    if quit:
        try:
            xl.Quit()
            del xl
        except:
            pass

    return attrs


def get_file_details(directory, filenames=None, open_excel=False):
    """Collect the a file or list of files attributes.

    Args:
        directory (str): Directory or filename to get attributes for
        filenames (str/list/tuple): If the given directory is a directory then a filename or list of files must be given
        open_excel (bool)[False]: If True read the properties by opening the excel file using the com object (slow).

    Returns:
         file_attrs (dict): Dictionary of {filename: {attribute_name: value}} or dictionary of {attribute_name: value}
            if a single file is given.
    """
    if os.path.isfile(directory):
        directory, filenames = os.path.dirname(directory), [os.path.basename(directory)]
    elif filenames is None:
        filenames = os.listdir(directory)
    elif not isinstance(filenames, (list, tuple)):
        filenames = [filenames]

    if not os.path.exists(directory):
        raise ValueError('The given directory does not exist!')

    # Open the com object
    sh = win32com.client.gencache.EnsureDispatch('Shell.Application', 0)  # Generates local compiled with make.py
    ns = sh.NameSpace(os.path.abspath(directory))

    # Get the directory file attribute column names
    cols = {}
    for i in range(512):  # 308 seemed to be max for excel file
        attrname = ns.GetDetailsOf(None, i)
        if attrname:
            cols[i] = attrname

    # Get the information for the files.
    files = {}
    for file in filenames:
        item = ns.ParseName(os.path.basename(file))
        files[os.path.abspath(item.Path)] = attrs = {}  # Store attributes in dictionary

        # Save attributes
        for i, attrname in cols.items():
            attrs[attrname] = ns.GetDetailsOf(item, i)

        # For xls file save special properties
        if os.path.splitext(file)[-1] == '.xls' or open_excel:
            try:
                xls_attrs = get_xls_properties(item.Path)
                attrs.update(xls_attrs)
            except (AttributeError, RuntimeError, Exception):
                pass

    # Clean up the com object
    try:
        sh.Quit()
    except:
        pass
    try:
        del sh
    except:
        pass

    if len(files) == 1:
        return files[list(files.keys())[0]]
    return files


if __name__ == '__main__':
    import argparse

    P = argparse.ArgumentParser(description="Read and print file details.")
    P.add_argument('filename', type=str, help='Filename to read and print the details for.')
    P.add_argument('-v', '--show-empty', action='store_true', help='If given print keys with empty values.')
    P.add_argument('-e', '--open-excel', action='store_true', help='If given open excel to read the attributes (slow).')
    ARGS = P.parse_args()

    # Argparse Variables
    FILENAME = ARGS.filename
    SHOW_EMPTY = ARGS.show_empty
    OPEN_EXCEL = ARGS.open_excel
    DETAILS = get_file_details(FILENAME, open_excel=OPEN_EXCEL)

    print(os.path.abspath(FILENAME))
    for k, v in DETAILS.items():
        if v or SHOW_EMPTY:
            print('\t', k, '=', v)
