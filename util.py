import json

COMMA = ","
EMPTY_STRING = ""
NEWLINE = "\n"

class ClaimsDataHeader:
    def __init__(self, line: str, ) -> None:
        self.fields = line.split(COMMA)

class ClaimsData:
    def __init__(self, data_header: ClaimsDataHeader, line: str) -> None:
        self.data_header = data_header
        self.data = line.split(COMMA)

def get_field_map():
    data = None
    with open('field_map.json') as f:
        # Load the contents of the file into a Python object
        data = json.load(f)

    return data

def read_claims_data():
    result = []
    with open('test.csv', 'r') as f:
        first = True
        data_header = None
        # Loop over each line in the file
        for line in f:
            line = line.replace(NEWLINE, EMPTY_STRING)
            if first:
                data_header = ClaimsDataHeader(line)
                first = False
            else:
                claim = ClaimsData(data_header, line)
                result.append(claim)

    return result

def write_file(file: str, data):
    _write_file_data(file, data, "w")

def _write_file_data(file: str, data, method: str):
    f = open(file,method)
    f.write(data)
    f.close()